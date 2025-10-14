import fs from "fs";
import { createRequire } from "module";

const require = createRequire(import.meta.url);

export function loadBootstrapCss() {
  const cssPath = require.resolve("bootstrap/dist/css/bootstrap.min.css");
  return fs.readFileSync(cssPath, "utf8");
}

export function loadBootstrapBundleJs() {
  // Хэрэв Bootstrap JS (tooltip, modal гэх мэт) хэрэгтэй бол:
  const jsPath = require.resolve("bootstrap/dist/js/bootstrap.bundle.min.js");
  return fs.readFileSync(jsPath, "utf8");
}

/** <head> хэсэгт оруулах HTML бэлдэнэ */
export function bootstrapHead({ withJs = false } = {}) {
  const css = loadBootstrapCss();
  const js = withJs ? `\n<script>${loadBootstrapBundleJs()}</script>` : "";
  return `<style id="bootstrap-css">${css}</style>${js}`;
}
