#target photoshop
app.displayDialogs = DialogModes.NO;

/** 設定 */
var OWNER_NAME = "松原";
var OUTPUT_PREFIX = "加工済 ";
var JPEG_QUALITY = 10;
var APPEND_LAYER_INDEX = false;

/** util */
function sanitizeName(s){ return s.replace(/[\\\/:\*\?"<>\|]/g,"_").replace(/^\s+|\s+$/g,""); }
function nextAvailableFile(folder, base, ext){
  var f = new File(folder.fsName + "/" + base + ext);
  if(!f.exists) return f;
  var i=1; for(;;){ f = new File(folder.fsName + "/" + base + "-" + i + ext); if(!f.exists) return f; i++; }
}

// --- 追加：一時プレフィックスを除去（layers_/layers-、tmp/__tmp-系 など） ---
function dropTempPrefixes(s){
  s = s.replace(/^\s*layers[_-]/i, "");
  s = s.replace(/^\s*(?:__?tmp(?:[_-][0-9a-f]+)?[_-]?)+/i, "");
  return s;
}
// パス末尾のセグメント名だけ取り出す
function leafOf(pathOrName){
  if (pathOrName.indexOf("/")>=0 || pathOrName.indexOf("\\")>=0){
    return pathOrName.replace(/\\/g,"/").replace(/\/+$/,"").split("/").pop();
  }
  return pathOrName;
}
// Unicodeの空白を広めに対応
function isSpace(code){
  if (code===0x20 || code===0x00A0 || code===0x1680 || code===0x202F || code===0x205F || code===0x3000) return true; // 半角/NBSP/OGHAM/ナローNBSP/MMSP/全角
  return (code>=0x2000 && code<=0x200A); // EN QUAD～HAIR SPACE
}
function rtrimSpaces(s){
  while (s.length && isSpace(s.charCodeAt(s.length-1))) s = s.substring(0, s.length-1);
  return s;
}
// 追加：安全にURLデコード（%xx をUnicodeへ / '+' をスペース扱い）
function urlDecodeSafe(s){
  var t = s;
  try { t = decodeURIComponent(t); } catch(e1){ try { t = decodeURI(t); } catch(e2){} }
  return t.replace(/\+/g, " ");
}

// 置き換え：最後のスペース以降を返す（URLデコード＋空白正規化）
function tokenAfterLastSpace(nameOrPath){
  // 末尾セグメント & 一時プレフィックス除去（layers_, __tmp_ など）
  var leaf = dropTempPrefixes(leafOf(nameOrPath));

  // ★URLエンコードを解除（%20 → 空白、%E5… → 日本語）
  leaf = urlDecodeSafe(leaf);

  // あらゆる空白を半角スペース1個に正規化（半角/全角/NBSP/2000-200A/202F/205F/タブ/改行）
  leaf = leaf.replace(/[ \u00A0\u1680\u2000-\u200A\u202F\u205F\u3000\t\r\n]+/g, " ").replace(/^ +| +$/g, "");

  // 最後のスペース以降を採用（なければ全体）
  var pos = leaf.lastIndexOf(" ");
  var base = (pos >= 0 && pos < leaf.length - 1) ? leaf.substring(pos + 1) : leaf;

  // 禁止記号のみ置換（空白は保持）
  return base.replace(/[\\\/:\*\?"<>\|]/g, "_").replace(/^ +| +$/g, "");
}

function collectAllLeafLayers(doc){
  var arr=[]; (function walk(g){
    for (var i=0;i<g.layers.length;i++){
      var L=g.layers[i];
      if (L.typename==="LayerSet") walk(L); else arr.push(L);
    }
  })(doc); return arr;
}
function toRgb8(doc){
  if (doc.mode !== DocumentMode.RGB) doc.changeMode(ChangeMode.RGB);
  if (doc.bitsPerChannel !== BitsPerChannelType.EIGHT) doc.bitsPerChannel = BitsPerChannelType.EIGHT;
}

/** main */
(function(){
  if (!app.documents.length){ alert("PSDを開いてから実行してください"); return; }
  var src = app.activeDocument;

  // 保存済みか安全に判定
  var hasPath=false, dirFs=null;
  try{ if(src.path && src.path.fsName){ hasPath=true; dirFs=src.path.fsName; } }catch(_e){ hasPath=false; }

  // ★ベースネーム：保存済み→親フォルダ名の「末尾空白以降」／未保存→ドキュメント名の「末尾空白以降」
  var baseName = hasPath
    ? tokenAfterLastSpace(dirFs)
    : tokenAfterLastSpace(src.name.replace(/\.[^\.]+$/,""));

  // 出力先
  var rootFs = hasPath ? dirFs : Folder.desktop.fsName;
  var outRoot = new Folder(rootFs + "/" + sanitizeName(OWNER_NAME + OUTPUT_PREFIX + baseName));
  if (!outRoot.exists) { try{ outRoot.create(); }catch(e){} }

  var layers = collectAllLeafLayers(src);
  if (layers.length === 0){ alert("書き出し対象レイヤーがありません"); return; }
  var idxWidth = (""+layers.length).length;

  var ok=0, ng=0, errs=[];
  for (var i=0;i<layers.length;i++){
    var lyr = layers[i];
    var layerName = sanitizeName(lyr.name || ("Layer"+(i+1)));
    var base = APPEND_LAYER_INDEX ? (("000000"+(i+1)).slice(-idxWidth)+"_"+layerName) : layerName;

    var ndoc=null;
    try{
      ndoc = app.documents.add(src.width, src.height, src.resolution, layerName,
                               NewDocumentMode.RGB, DocumentFill.TRANSPARENT);
      app.activeDocument = src; // 元PSDを前面に戻す
    }catch(eNew){
      ng++; errs.push("新規作成失敗: "+layerName+" / "+eNew);
      continue;
    }
    try{
      lyr.duplicate(ndoc, ElementPlacement.PLACEATBEGINNING);
    }catch(eDup){
      ng++; errs.push("複製失敗: "+layerName+" / "+eDup);
      try{ ndoc.close(SaveOptions.DONOTSAVECHANGES); }catch(_e){}
      continue;
    }

    app.activeDocument = ndoc;
    try{
      toRgb8(ndoc);
      ndoc.flatten();
      var outFile = nextAvailableFile(outRoot, base, ".jpg");
      var opt = new JPEGSaveOptions();
      opt.quality = JPEG_QUALITY;
      opt.embedColorProfile = true;
      opt.formatOptions = FormatOptions.OPTIMIZEDBASELINE;
      ndoc.saveAs(outFile, opt, true, Extension.LOWERCASE);
      ok++;
    }catch(eSave){
      ng++; errs.push("保存失敗: "+layerName+" / "+eSave);
    }finally{
      try{ ndoc.close(SaveOptions.DONOTSAVECHANGES); }catch(_e){}
    }
  }

  var msg = "完了: " + outRoot.fsName + "\n保存: " + ok + " 件 / 失敗: " + ng + " 件";
  if (errs.length) msg += "\n--- 失敗詳細(最大10件) ---\n" + errs.slice(0,10).join("\n");
  alert(msg);
})();
