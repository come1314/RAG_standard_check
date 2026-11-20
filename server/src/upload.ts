import fs from "node:fs/promises";
import path from "node:path";
import { PATHS, SUPPORTED_EXTENSIONS, isSupportedFile } from "./config.js";

export interface UploadFailure {
  name: string;
  reason: string;
}

export interface UploadOutcome {
  saved: string[];
  failures: UploadFailure[];
}

/**
 * The slice of multer's file object we actually use. Declared structurally so
 * this module doesn't depend on multer's global `Express.Multer` namespace —
 * tsconfig restricts automatic @types inclusion to `node`, and reaching for
 * that global here would fail to compile.
 */
export interface UploadedFile {
  originalname: string;
  buffer: Buffer;
}

const MAX_NAME_LENGTH = 180;

// Control characters plus the set Windows forbids in file names. Deliberately a
// \u0000-\u001f range rather than a bare dash: space is 0x20 and digits are
// 0x30-0x39, so a name like "文档 200 104-2026 xxx.docx" survives untouched.
const FORBIDDEN_CHARS = /[\u0000-\u001f\u007f<>:"/\|?*]/g;

/**
 * Turn an attacker-controlled filename into something safe to place inside
 * `knowledgeFile/`.
 *
 * Separators are stripped rather than merely rejected, so a name that walks up
 * the directory tree can never escape, and leading dots are refused so a bare
 * ".docx" or a hidden temp file doesn't end up in the knowledge base.
 */
export function sanitizeFileName(original: string): { name: string } | { error: string } {
  const withoutSeparators = original.split("\\").join("/");
  const base = withoutSeparators.split("/").pop() ?? "";
  const cleaned = base.replace(FORBIDDEN_CHARS, "").trim();

  if (!cleaned || cleaned === "." || cleaned === "..") {
    return { error: "文件名无效" };
  }
  if (cleaned.startsWith(".")) {
    return { error: "文件名不能以点开头" };
  }
  if (cleaned.length > MAX_NAME_LENGTH) {
    return { error: `文件名过长（上限 ${MAX_NAME_LENGTH} 字符）` };
  }
  if (!isSupportedFile(cleaned)) {
    return { error: `仅支持 ${SUPPORTED_EXTENSIONS.join(" / ")} 文件` };
  }
  return { name: cleaned };
}

/** A .docx is a ZIP package — this rejects a renamed .exe or .txt instantly. */
function looksLikeDocx(buffer: Buffer): boolean {
  return (
    buffer.length > 4 &&
    buffer[0] === 0x50 &&
    buffer[1] === 0x4b &&
    buffer[2] === 0x03 &&
    buffer[3] === 0x04
  );
}

async function exists(target: string): Promise<boolean> {
  try {
    await fs.access(target);
    return true;
  } catch {
    return false;
  }
}

/**
 * Validate and write uploaded buffers into `knowledgeFile/`.
 *
 * Existing files are never overwritten: these are the source documents, and a
 * silent replace would destroy one with no undo. A name collision is reported
 * as a per-file failure instead, so the rest of the batch still lands.
 */
export async function saveUploads(files: UploadedFile[]): Promise<UploadOutcome> {
  const saved: string[] = [];
  const failures: UploadFailure[] = [];

  await fs.mkdir(PATHS.knowledgeDir, { recursive: true });

  for (const file of files) {
    const check = sanitizeFileName(file.originalname);
    if ("error" in check) {
      failures.push({ name: file.originalname, reason: check.error });
      continue;
    }

    if (!looksLikeDocx(file.buffer)) {
      failures.push({
        name: file.originalname,
        reason: "内容不是有效的 .docx（文件可能损坏，或只是改了扩展名）",
      });
      continue;
    }

    const target = path.join(PATHS.knowledgeDir, check.name);
    if (await exists(target)) {
      failures.push({ name: check.name, reason: "同名文件已存在，未覆盖" });
      continue;
    }

    try {
      await fs.writeFile(target, file.buffer);
      saved.push(check.name);
    } catch (cause) {
      failures.push({
        name: check.name,
        reason: cause instanceof Error ? cause.message : String(cause),
      });
    }
  }

  return { saved, failures };
}
