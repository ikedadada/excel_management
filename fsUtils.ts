import fs from "fs";
import path from "path";

export function walk(dir: string, onFile: (file: string) => void): void {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full, onFile);
    else if (entry.isFile() && full.endsWith(".xlsx")) {
      onFile(full);
    }
  }
}

export function deleteDir(dir: string): void {
  if (fs.existsSync(dir)) {
    fs.rmSync(dir, { recursive: true, force: true });
    console.log(`Deleted directory: ${dir}`);
  }
}
