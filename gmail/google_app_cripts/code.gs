ELF_SS = SpreadsheetApp.getActiveSpreadsheet();
var SS_ID = SELF_SS.getId();
var CREDS = NLCUTIL_load_creds()

NB_CLFS = 3

var CONF_INDEX = {
    ws_name: 0,
    start_col: 1,
    start_row: 2,
    train_column: 3,
    intent1_col: 4,
    result1_col: 5,
    restime1_col: 6,
    intent2_col: 7,
    result2_col: 8,
    restime2_col: 9,
    intent3_col: 10,
    result3_col: 11,
    restime3_col: 12,
    log_ws: 13,
    notif_opt: 14,
    notif_ws: 15,
    query: 16,
}

CONFIG_SET = {
    ss_id: SS_ID,
    ws_name: '設定',
    st_start_row: 2,
    st_start_col: 2,
    exc_start_row: 3,
    exc_start_col: 5,
    notif_start_row: 2,
    notif_start_col: 1,
    notif_from: "takahiko.naruse@g.softbank.co.jp",
    notif_sender: "NLC Classsifier",
    result_override: false,
    clfs_start_col: 8,
    clfs_start_row: 3,
    log_start_col: 1,
    log_start_row: 2,
}


function onOpen() {

  var ui = SpreadsheetApp.getUi();
    ui.createMenu('Watson')
      .addItem('メール取得', 'MAILUTIL_load_messages')
      .addItem('分類', 'MAILUTIL_classify_all')
      .addItem('学習', 'MAILUTIL_train_all')
      .addToUi();
      
  NLCUTIL_exec_check_clfs();
}



