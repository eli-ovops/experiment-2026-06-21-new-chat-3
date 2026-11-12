import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/outputs/长空之王QC小组_QCC成果报告_提高55-72黑丝优等率_视觉重做版.pptx";
const TMP = "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/work/presentations/qcc_55_72_heisi/tmp";
const OUT = `${TMP}/final_preview_v2`;
const LAYOUT = `${TMP}/final_layout_v2`;

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  await fs.mkdir(OUT, { recursive: true });
  await fs.mkdir(LAYOUT, { recursive: true });
  const presentation = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  for (const [idx, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(idx + 1).padStart(2, "0")}`;
    await writeBlob(`${OUT}/${stem}.png`, await presentation.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${LAYOUT}/${stem}.json`, await layout.text());
  }
  await writeBlob(`${OUT}/deck-montage.webp`, await presentation.export({ format: "webp", montage: true, scale: 1 }));
  const inspect = await presentation.inspect({ kind: "slide,textbox,shape,table,chart,image", maxChars: 12000 });
  await fs.writeFile(`${TMP}/qa/final-inspect-v2.ndjson`, inspect.ndjson);
  console.log(`slides=${presentation.slides.items.length}`);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
