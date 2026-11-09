const BASE = 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/';
let py;
let chain = Promise.resolve();
async function handle({id,type,payload}) {
  try {
    let result;
    if(type==='init') {
      if(!py) { // Python 환경은 처음 한 번만 준비 (분류를 바꿀 때는 다시 준비하지 않음)
        postMessage({type:'status',text:'Python 실행 환경 준비 중… (처음에만 10~30초)'});
        importScripts(BASE+'pyodide.js');
        py=await loadPyodide({indexURL:BASE});
        await py.loadPackage(['numpy','micropip']);
        await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl==3.1.5')");
      }
      const response=await fetch('./model.py',{cache:'no-store'});
      if(!response.ok) throw Error('model.py를 불러올 수 없습니다.');
      await py.runPythonAsync(await response.text());
      postMessage({type:'status',text:'Excel 읽기와 계수 검증 중…'});
      py.globals.set('raw_input',new Uint8Array(payload.buffer));
      py.globals.set('level_input',payload.level);
      await py.runPythonAsync('raw_bytes = raw_input.to_py()');
      result=await py.runPythonAsync('initialize(raw_bytes, level_input)');
      await py.runPythonAsync('del raw_input, raw_bytes');
    } else if(type==='explore') {
      py.globals.set('sheet_input',payload.sheet);
      py.globals.set('row_input',payload.row||'');
      result=py.runPython('explore(sheet_input, row_input)');
    } else if(type==='linkage') {
      result=py.runPython('linkage()');
    } else {
      py.globals.set('scenario_input',JSON.stringify(payload));
      result=py.runPython('simulate(scenario_input)');
    }
    postMessage({id,ok:true,result:JSON.parse(result)});
  } catch(error) { postMessage({id,ok:false,error:String(error)}); }
}
onmessage=e=>{ chain=chain.then(()=>handle(e.data)); };
