// worker.js - Pyodide Web Worker for Zero-Backend OOXML Manipulation
importScripts("https://cdn.jsdelivr.net/pyodide/v0.27.2/full/pyodide.js");

let pyodide = null;
let readyPromise = null;

async function initEngine() {
  self.postMessage({ type: "status", stage: "loading", text: "1/3 載入 Pyodide WASM 引擎..." });
  pyodide = await loadPyodide();

  self.postMessage({ type: "status", stage: "loading", text: "2/3 載入 lxml 與安裝 python-docx..." });
  await pyodide.loadPackage(["lxml", "micropip"]);
  const micropip = pyodide.pyimport("micropip");
  await micropip.install("python-docx");

  self.postMessage({ type: "status", stage: "loading", text: "3/3 掛載 xml_engine.py 至 MEMFS..." });
  const engineResp = await fetch("./xml_engine.py");
  if (!engineResp.ok) {
    throw new Error(`無法讀取 xml_engine.py (HTTP ${engineResp.status})`);
  }
  const engineCode = await engineResp.text();

  pyodide.FS.mkdirTree("/app");
  pyodide.FS.mkdirTree("/work");
  pyodide.FS.writeFile("/app/xml_engine.py", engineCode);

  await pyodide.runPythonAsync(`
import sys, json
sys.path.insert(0, "/app")
import xml_engine
`);

  self.postMessage({ type: "status", stage: "ready", text: "Pyodide 引擎就緒" });
}

readyPromise = initEngine().catch((err) => {
  self.postMessage({ type: "status", stage: "error", text: `初始化失敗：${err.message}` });
  throw err;
});

self.onmessage = async (event) => {
  const { id, action, payload } = event.data;
  try {
    await readyPromise;

    if (action === "load_file") {
      const bytes = new Uint8Array(payload.buffer);
      pyodide.FS.writeFile("/work/template.docx", bytes);
      pyodide.FS.writeFile("/work/current.docx", bytes);
      const summary = await pyodide.runPythonAsync(`
xml_engine.inspect_docx("/work/current.docx", max_items=600)
`);
      self.postMessage({ id, ok: true, result: summary });
      return;
    }

    if (action === "reset_template") {
      const orig = pyodide.FS.readFile("/work/template.docx");
      pyodide.FS.writeFile("/work/current.docx", orig);
      const copyBuf = orig.slice().buffer;
      const summary = await pyodide.runPythonAsync(`
xml_engine.inspect_docx("/work/current.docx", max_items=600)
`);
      self.postMessage({ id, ok: true, result: summary, fileBuffer: copyBuf }, [copyBuf]);
      return;
    }

    if (action === "docx_inspect") {
      pyodide.globals.set("_arg_empty", Boolean(payload.include_empty));
      pyodide.globals.set("_arg_runs", Boolean(payload.include_runs));
      pyodide.globals.set("_arg_prefix", payload.target_prefix || null);
      pyodide.globals.set("_arg_max", Number(payload.max_items || 300));
      const output = await pyodide.runPythonAsync(`
xml_engine.inspect_docx(
    "/work/current.docx",
    include_empty=_arg_empty,
    include_runs=_arg_runs,
    target_prefix=_arg_prefix,
    max_items=_arg_max,
)
`);
      self.postMessage({ id, ok: true, result: output });
      return;
    }

    if (action === "docx_apply_plan") {
      pyodide.globals.set("_arg_plan_json", JSON.stringify(payload.plan || {}));
      const report = await pyodide.runPythonAsync(`
plan = json.loads(_arg_plan_json)
res = xml_engine.apply_plan_docx("/work/current.docx", plan=plan, output_path="/work/current.docx")
res.to_compact_report()
`);
      const updatedBytes = pyodide.FS.readFile("/work/current.docx");
      const copyBuf = updatedBytes.slice().buffer;
      self.postMessage({ id, ok: true, result: report, fileBuffer: copyBuf }, [copyBuf]);
      return;
    }

    if (action === "docx_batch_edit") {
      pyodide.globals.set("_arg_edits_json", JSON.stringify(payload.edits || []));
      const report = await pyodide.runPythonAsync(`
edits = json.loads(_arg_edits_json)
res = xml_engine.batch_edit_docx("/work/current.docx", edits=edits, output_path="/work/current.docx")
res.to_compact_report()
`);
      const updatedBytes = pyodide.FS.readFile("/work/current.docx");
      const copyBuf = updatedBytes.slice().buffer;
      self.postMessage({ id, ok: true, result: report, fileBuffer: copyBuf }, [copyBuf]);
      return;
    }

    if (action === "docx_clone_node") {
      pyodide.globals.set("_arg_target", String(payload.target));
      pyodide.globals.set("_arg_anchor", String(payload.anchor_text));
      pyodide.globals.set("_arg_items_json", JSON.stringify(payload.items || []));
      pyodide.globals.set("_arg_mode", String(payload.mode || "after"));
      const report = await pyodide.runPythonAsync(`
items = json.loads(_arg_items_json)
res = xml_engine.clone_node_docx(
    "/work/current.docx",
    target=_arg_target,
    anchor_text=_arg_anchor,
    items=items,
    mode=_arg_mode,
    output_path="/work/current.docx",
)
res.to_compact_report()
`);
      const updatedBytes = pyodide.FS.readFile("/work/current.docx");
      const copyBuf = updatedBytes.slice().buffer;
      self.postMessage({ id, ok: true, result: report, fileBuffer: copyBuf }, [copyBuf]);
      return;
    }

    throw new Error(`Unknown worker action: ${action}`);
  } catch (err) {
    self.postMessage({ id, ok: false, error: err.message || String(err) });
  }
};
