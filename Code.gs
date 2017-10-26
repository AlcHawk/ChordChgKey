// === Create sidebar in HTML on spreadsheet UI ===
function createSidebarHTML() {
  var ui = HtmlService.createHtmlOutputFromFile("ChordPanel").setTitle("Chord Panel");
  DocumentApp.getUi().showSidebar(ui);
}

// === Open sidebar HTML ===
function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
      .addItem("ChordPanel", "createSidebarHTML")
      .addToUi();
}

// === Custom functions ===
function extractContentToArray() {
  var doc = DocumentApp.getActiveDocument();
  var con = doc.getBody().getText();
  if (con.indexOf("\r") >= 0 ) {
    var con_array = con.split("\r");
    Logger.log("Use return")
  } else if (con.indexOf("\n") >= 0 ) {
    var con_array = con.split("\n");
    Logger.log("Use new line")
  }
  
//  Logger.log(con_array);
  return(con_array);
}

function extractKey() {
  var doc = DocumentApp.getActiveDocument();
  var con = doc.getBody().getText();

  var key_line = con.match(/[Kk][Ee][Yy]\:[ A-Za-z]+/);
  var key = key_line[0].replace(/\s/g,"").split(":")[1];

  return(key.toUpperCase());
}

function chordChangeByKey(chord, org_key, chg_key) {
  key_map = {
    "C": 1,
    "D": 2,
    "E": 3,
    "F": 4,
    "G": 5,
    "A": 6,
    "B": 7,
  }
  if (chord != "") {
    key_dist = key_map[chg_key] - key_map[org_key];
    chord_name = chord.substr(0,1);
    change_chord_num = ( key_map[chord_name] + key_dist - 1 ) % 7;
    if (change_chord_num < 0) {
      change_chord_num = change_chord_num + 7
    }
    change_chord = Object.keys(key_map)[change_chord_num];
    new_chord = change_chord + chord.substr(1, chord.length - 1);
  } else {
    new_chord = ""
  }
  return(new_chord);
}

function replaceChord(new_key) {
  var con_ary = extractContentToArray();
  var key = extractKey();
  
  for (var i = 0; i < con_ary.length; i++) {
    if (con_ary[i].replace(/\s/, "").substr(0, 3).toUpperCase() === "KEY") {
      var key_append = "Key: "+new_key;
      con_ary[i] = key_append;
    } else if (con_ary[i].trim() != "" && con_ary[i].replace(/[A-Za-z0-9\s\n]+/, "") === "" ) {
      var chord_line = con_ary[i].split(" ");
      var chord_line_re = chord_line.map(function(x) { return chordChangeByKey(x, key, new_key); });
      var chord_re = chord_line_re.join(" ");
      con_ary[i] = chord_re;
    }
  }
//  Logger.log(con_ary);
  return(con_ary)
}

function newContentReplace(new_key) {
  var doc_body = DocumentApp.getActiveDocument().getBody();
  
//  var old_con = doc_body.getText();
//  Logger.log(old_con);
  var new_con = replaceChord(new_key).join("\n");
  Logger.log(new_con);

  doc_body.setText(new_con);
}

function outputToPDF() {
}

