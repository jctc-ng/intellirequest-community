import {build} from "esbuild";
import {mkdir,copyFile,writeFile} from "node:fs/promises";
await mkdir("dist/assets",{recursive:true});
await copyFile("index.html","dist/index.html");
await copyFile("assets/nfiu-favicon.svg","dist/assets/nfiu-favicon.svg");
await writeFile("dist/.nojekyll","");
await build({entryPoints:["src/app.js"],outfile:"dist/assets/app.js",
  bundle:true,minify:true,format:"iife",target:["es2022"],platform:"browser"});
console.log("Static site prepared under dist/. No document processing server is involved.");
