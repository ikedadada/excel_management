import { convertBook } from "./excelUtils";
import { walk, deleteDir } from "./fsUtils";

const SRC = process.argv[2] ?? "docs";
const DST = process.argv[3] ?? "diff";

deleteDir(DST);
walk(SRC, (file) => {
  void convertBook(file, DST);
});
