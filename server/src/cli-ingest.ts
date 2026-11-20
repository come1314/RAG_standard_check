import { ingest, listKnowledgeFiles, readIndexMeta } from "./ingest.js";
import { PATHS, RAG } from "./config.js";

/**
 * Build the knowledge index from the command line:
 *
 *   npm run ingest --prefix server
 *
 * Useful for pre-warming the vector store (and downloading the embedding model)
 * before starting the API server.
 */
async function main(): Promise<void> {
  const files = await listKnowledgeFiles();
  console.log(`knowledge dir : ${PATHS.knowledgeDir}`);
  console.log(`files found   : ${files.length}`);
  console.log(`embedding     : ${RAG.embeddingModel}`);
  console.log(`vector store  : ${PATHS.vectorDir}`);

  if (files.length === 0) {
    console.error("\nNo .docx files found — nothing to index.");
    process.exitCode = 1;
    return;
  }

  const result = await ingest();
  const meta = await readIndexMeta();

  console.log("\nIndexed:");
  for (const file of result.files) console.log(`  · ${file}`);
  console.log(`\n${meta?.chunkCount ?? result.chunkCount} chunks written to ${PATHS.vectorDir}`);
  console.log(`Built at ${result.builtAt}`);
}

main().catch((error: unknown) => {
  console.error("\nIngest failed:", error instanceof Error ? error.message : error);
  process.exit(1);
});
