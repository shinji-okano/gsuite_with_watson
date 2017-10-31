// The MIT License (MIT)
//
// Copyright (c) 2017 SoftBank Group Corp.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.





// ----------------------------------------------------------------------------
// グローバル変

NB_FIELDS = 5;
MAIL_FIELDS = {
    ID: 0,
    SUBJECT: 1,
    TO: 2,
    BODY: 3,
    NORM_BODY: 4
};
TRAIN_COLUMN = {
    SUBJECT: "件名のみ",
    BODY: "本文のみ",
    BOTH: "件名・本文両方",
};

function MAILUTIL_load_config(config_set) {

    var ss = SpreadsheetApp.openById(config_set.ss_id);

    var sheet = ss.getSheetByName(config_set.ws_name);

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var conf_list = sheet.getRange(config_set.st_start_row, config_set.st_start_col, lastRow, 1).getValues()

    var sheet_conf = {}
    sheet_conf["ws_name"] = conf_list[CONF_INDEX.ws_name][0]
    sheet_conf["start_col"] = conf_list[CONF_INDEX.start_col][0]
    sheet_conf["start_row"] = conf_list[CONF_INDEX.start_row][0]
    sheet_conf["train_column"] = conf_list[CONF_INDEX.train_column][0]
    sheet_conf["intent_col"] = [conf_list[CONF_INDEX.intent1_col][0], conf_list[CONF_INDEX.intent2_col][0], conf_list[CONF_INDEX.intent3_col][0]]
    sheet_conf["result_col"] = [conf_list[CONF_INDEX.result1_col][0], conf_list[CONF_INDEX.result2_col][0], conf_list[CONF_INDEX.result3_col][0]]
    sheet_conf["restime_col"] = [conf_list[CONF_INDEX.restime1_col][0], conf_list[CONF_INDEX.restime2_col][0], conf_list[CONF_INDEX.restime3_col][0]]
    sheet_conf["log_ws"] = conf_list[CONF_INDEX.log_ws][0]
    sheet_conf["query"] = conf_list[CONF_INDEX.query][0]

    var notif_conf = {};
    notif_conf["option"] = conf_list[CONF_INDEX.notif_opt][0];
    notif_conf["ws_name"] = conf_list[CONF_INDEX.notif_ws][0];

    var exc_list = sheet.getRange(config_set.exc_start_row, config_set.exc_start_col, lastRow - config_set.exc_start_row + 1, 1).getValues()
    var exc_conf = {}
    var re_list = []
    for (var i = 0; i < exc_list.length; i++) {
        var exc_re = exc_list[i][0]
        if (exc_re != "") {
            re_list.push(exc_re)
        }
    }
    exc_conf["re_list"] = re_list

    return {
        sheet_conf: sheet_conf,
        notif_conf: notif_conf,
        exc_conf: exc_conf,
    }
}



function MAILUTIL_load_messages() {

  var conf = MAILUTIL_load_config( CONFIG_SET )

  var mail_set = {
    ss_id: SS_ID,
    ws_name: conf.sheet_conf.ws_name,
    start_col: conf.sheet_conf.start_col,
    start_row: conf.sheet_conf.start_row,
    query: conf.sheet_conf.query,
    exc_res: conf.exc_conf.re_list,
    msgs: []
  }

  var msgs = MAILUTIL_get_messages( mail_set )

  mail_set.msgs = msgs

  mail_set = MAILUTIL_trim_exc( mail_set )

  MAILUTIL_update_data( mail_set )
}



function MAILUTIL_get_messages(mail_set) {

    Logger.log("### MAILUTIL_get_messages")

    var result = []
    var threads = GmailApp.search(mail_set.query, 0, 10);
    // in:inbox is:read

    Logger.log("THREADS:" + threads.length)
    for (var i = 0; i < threads.length; i++) {

        var thread = threads[i];
        var msgs = thread.getMessages();

        for (var j = 0; j < msgs.length; j++) {
            var msg = msgs[j];
            //valMsgs[j][0] = msg.getDate()
            //valMsgs[j][1] = msg.getFrom()

            var res_msg = [msg.getId(), msg.getSubject(), msg.getTo(), msg.getPlainBody()]
            result.push(res_msg)
        }

    }

    return result
}


function MAILUTIL_trim_exc(mail_set) {

    re = "\[image\:.*\]"
    regexp = new RegExp(re, 'g')

    //regexp = "/このメールアドレスは送信専用の為/g"

    for (var i = 0; i < mail_set.msgs.length; i++) {

        var msg = mail_set.msgs[i]
        var buf = msg[MAIL_FIELDS.BODY]

        for (var j = 0; j < mail_set.exc_res.length; j++) {

            Logger.log(mail_set.exc_res[j])

            //mail_set.exc_res[j].replace(/\\/g, '\\\\')
            //Logger.log( mail_set.exc_res[j] )

            var regexp = new RegExp(mail_set.exc_res[j], 'g')
            //var regexp = new RegExp(  mail_set.exc_res[j].replace(/[\\^$.*+?()[\]{}|]/g, '\\$&'));

            buf = buf.replace(regexp, "")

        }

        mail_set.msgs[i][MAIL_FIELDS.NORM_BODY] = buf.trim()

    }

    return mail_set

}



function MAILUTIL_update_data(mail_set) {

    var ss = SpreadsheetApp.openById(mail_set.ss_id);

    var sheet = ss.getSheetByName(mail_set.ws_name);

    if (sheet == null) {
        sheet = ss.insertSheet(mail_set.ws_name)
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    //sheet.getRange( start_row, start_col, lastRow-start_row+1, 4 ).setBackgroundRGB(128, 0, 0 )

    Logger.log("LASTROW" + lastRow)
    if (lastRow >= mail_set.start_row) {

        var entries = sheet.getRange(mail_set.start_row, mail_set.start_col, lastRow - mail_set.start_row + 1, NB_FIELDS).getValues();

    } else {
        var entries = []
        for (cnt = 1; cnt < mail_set.start_row - lastRow; cnt++) {
            sheet.appendRow(["", "", "", "", ""])
        }
        lastRow = sheet.getLastRow();
    }


    var row_cnt = 0
    for (var i = 0; i < mail_set.msgs.length; i++) {

        var isMatch = 0
        for (var j = 0; j < entries.length; j++) {

            if (entries[j][MAIL_FIELDS.ID] == mail_set.msgs[i][MAIL_FIELDS.ID]) {
                isMatch = 1
                break
            }
        }

        if (isMatch == 0) {
            var record = [mail_set.msgs[i][MAIL_FIELDS.ID],
                mail_set.msgs[i][MAIL_FIELDS.SUBJECT],
                mail_set.msgs[i][MAIL_FIELDS.TO],
                mail_set.msgs[i][MAIL_FIELDS.BODY],
                mail_set.msgs[i][MAIL_FIELDS.NORM_BODY],

            ]

            sheet.getRange(lastRow + row_cnt + 1, mail_set.start_col, 1, NB_FIELDS).setValues([record])
            //sheet.getRange( lastRow + row_cnt + 1, data_set.start_col, 1, NB_FIELDS ).setBackgroundRGB(255, 255, 0 )
            row_cnt += 1
        }
    }
}




function MAILUTIL_train_set(clf_no) {

    Logger.log("### MAILUTIL_train_set", clf_no);

    var conf = MAILUTIL_load_config(CONFIG_SET);

    var train_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_row: conf.sheet_conf.start_row,
        start_col: conf.sheet_conf.start_col,
        end_row: -1,
        train_column: conf.sheet_conf.train_column,
        text_col: conf.sheet_conf.train_column,
        class_col: conf.sheet_conf.intent_col[clf_no - 1],
        clf_no: clf_no,
        clf_name: CLFNAME_PREFIX + clf_no,
    };

    var train_result = MAILUTIL_train(train_set, CREDS.username, CREDS.password);

    var log_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.log_ws,
        start_col: CONFIG_SET.log_start_col,
        start_row: CONFIG_SET.log_start_row
    };

    NLCUTIL_log_train(log_set, train_set, train_result);
}


function MAILUTIL_train(train_set, creds_username, creds_password) {

    Logger.log("### MAILUTIL_train");

    var clfs = NLCAPI_get_classifiers(creds_username, creds_password);
    if (clfs.status != 200) {
        return {
            status: clfs.code,
            description: clfs.status,
        };
    }

    
    var ss = SpreadsheetApp.openById(train_set.ss_id);

    var sheet = ss.getSheetByName(train_set.ws_name);

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries = sheet.getRange(train_set.start_row, train_set.start_col, lastRow - train_set.start_row + 1, lastCol - train_set.start_col).getValues();

    var row_cnt = 0;
    var csvString = "";

    for (var cnt = 0; cnt < entries.length; cnt++) {
    
        var class_name = NLCUTIL_norm_text(String(entries[cnt][train_set.class_col - train_set.start_col])).trim()
        if (class_name == "") continue

        var subject_text = NLCUTIL_norm_text(String(entries[cnt][MAIL_FIELDS.SUBJECT])).trim();
        var body_text = NLCUTIL_norm_text(String(entries[cnt][MAIL_FIELDS.NORM_BODY])).trim();

        var train_text;
        if (train_set.train_column == TRAIN_COLUMN.SUBJECT) {
            train_text = subject_text;
        } else if (train_set.train_column == TRAIN_COLUMN.BODY) {
            train_text = body_text;
        } else {
            if (subject_text == "") {
                train_text = body_text;
            } else if (body_text == "") {
                train_text = subject_text;
            } else {
                train_text = subject_text + " " + body_text;
            }
        }
        if (train_text.length == 0) continue;

        if (train_text.length > 1024) {
            train_text = train_text.substring(0, 1024);
        }

        csvString = csvString + '"' + train_text + '","' + class_name + '"' + "\r\n";
        row_cnt += 1;
    }


    var clf_info = NLCUTIL_clf_vers(clfs.body.classifiers, train_set.clf_name);

    var new_version = (clf_info.max_ver + 1);
    var clf_name = train_set.clf_name + "#__#" + new_version;

    var res = NLCAPI_post_classifiers(creds_username, creds_password, csvString, clf_name, 'ja');
    if (clf_info.count >= 2 && res.status == 200) {

        NLCAPI_delete_classifier(creds_username, creds_password, clf_info.clfs[clf_info.min_ver].classifier_id);

        //NLCUTIL_log_delete( clf_info.clfs[ clf_info.min_ver ] )

    }

    NLCUTIL_exec_check_clfs();
    NLCUTIL_set_trigger("NLCUTIL_exec_check_clfs", 1);

    var result = {
        status: res.status,
        nlc: res,
        rows: row_cnt,
        version: new_version
    };
    return result;
}

function MAILUTIL_train_all() {
    for (var i = 1; i <= NB_CLFS; i++) {
        MAILUTIL_train_set(i);
    }
}






function MAILUTIL_classify_all() {

    Logger.log("### MAILUTIL_classify_all");

    var conf = MAILUTIL_load_config(CONFIG_SET);

    var notif_conf = {
        ss_id: SS_ID,
        ws_name: conf.notif_conf.ws_name,
        start_col: CONFIG_SET.notif_start_col,
        start_row: CONFIG_SET.notif_start_row,
    };

    var rules = NLCUTI_load_notif_rules(notif_conf);

    var notif_set = {
        rules: rules,
        from: CONFIG_SET.notif_from,
        sender: CONFIG_SET.notif_sender,
        result_cols: conf.sheet_conf.result_col,
    };

    var log_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.log_ws,
        start_col: CONFIG_SET.log_start_col,
        start_row: CONFIG_SET.log_start_row
    };

    var test_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_col: conf.sheet_conf.start_col,
        start_row: conf.sheet_conf.start_row,
        end_row: -1,
        text_col: conf.sheet_conf.train_column,
        notif_set: notif_set,
    };

    var clf_ids = [];
    for (var i = 0; i < NB_CLFS; i++) {

        var clf_name = CLFNAME_PREFIX + String(i + 1);

        test_set.clf_no = i + 1;
        test_set.clf_name = clf_name;
        test_set.result_col = conf.sheet_conf.result_col[i];
        test_set.restime_col = conf.sheet_conf.restime_col[i];

        var clf = NLCUTIL_select_clf(clf_name, CREDS.username, CREDS.password);
        if (clf.status == "Training") {
            NLCUTIL_log_classify(log_set, test_set, {
                status: "900",
                description: "トレーニング中",
                clf_id: clf.clf_id
            });
        } else if (clf.status == "Nothing" ) {
                NLCUTIL_log_classify(log_set, test_set, {
                    status: "900",
                    description: "分類器なし",
                    clf_id: ""
                });
        } else if (clf.status != "Available") {
                NLCUTIL_log_classify(log_set, test_set, {
                    status: "800",
                    description: clf.status,
                    clf_id: ""
                });
        }

        clf_ids.push({
            id: clf.clf_id,
            status: clf.status
        });
    }

    var ss = SpreadsheetApp.openById(SS_ID);

    var sheet = ss.getSheetByName(conf.sheet_conf.ws_name);

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries = sheet.getRange(conf.sheet_conf.start_row, conf.sheet_conf.start_col, lastRow - conf.sheet_conf.start_row + 1, lastCol - conf.sheet_conf.start_col + 1).getValues();

    var hasError = 0;
    var row_cnt = 0;
    var nlc_res;
    var res_rows = [0, 0, 0];
    for (var cnt = 0; cnt < entries.length; cnt++) {

        var subject_text = NLCUTIL_norm_text(String(entries[cnt][MAIL_FIELDS.SUBJECT])).trim();
        var body_text = NLCUTIL_norm_text(String(entries[cnt][MAIL_FIELDS.NORM_BODY])).trim();

        var test_text;
        if (conf.train_column == TRAIN_COLUMN.SUBJECT) {
            test_text = subject_text;
        } else if (conf.train_column == TRAIN_COLUMN.BODY) {
            test_text = body_text;
        } else {
            if (subject_text == "") {
                test_text = body_text;
            } else if (body_text == "") {
                test_text = subject_text;
            } else {
                test_text = subject_text + " " + body_text;
            }
        }
        if (test_text.length == 0) continue;

        if (test_text.length > 1024) {
            test_text = test_text.substring(0, 1024);
        }

        var updates = 0;
        var upd_flg = [0, 0, 0];
        for (var j = 0; j < NB_CLFS; j++) {

            if (clf_ids[j].status != "Available") {
                continue;
            }

            var result_text;
            if (lastCol < conf.sheet_conf.result_col[j]) {
                result_text = "";
            } else {
                result_text = entries[cnt][conf.sheet_conf.result_col[j] - conf.sheet_conf.start_col];
            }

            if (result_text != "" && CONFIG_SET.result_override != true) continue;

            nlc_res = NLCAPI_post_classify(CREDS.username, CREDS.password, clf_ids[j].id, test_text);
            if (nlc_res.status != 200) {
                hasError = 1;
                err_res = nlc_res;
            } else {
                r = nlc_res.body.top_class;
                sheet.getRange(conf.sheet_conf.start_row + cnt, conf.sheet_conf.result_col[j], 1, 1).setValue(r);
                t = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
                sheet.getRange(conf.sheet_conf.start_row + cnt, conf.sheet_conf.restime_col[j], 1, 1).setValue(t);
                updates += 1;
                res_rows[j] += 1;
                upd_flg[j] = 1;
            }
        }

        if (updates > 0) {
            row_cnt += 1;
            var record = sheet.getRange(conf.sheet_conf.start_row + cnt, 1, 1, lastCol).getValues();
            NLCUTIL_check_notify(notif_set, record[0], upd_flg);
        }
    }

    for (var k = 0; k < NB_CLFS; k++) {

        if (clf_ids[k].status != "Available") continue;

        var clfname = CLFNAME_PREFIX + String(k + 1);
        test_set.clf_no = k + 1;
        test_set.clf_name = clfname;
        test_set.result_col = conf.sheet_conf.result_col[k];
        test_set.restime_col = conf.sheet_conf.restime_col[k];

        var test_result;
        if (res_rows[k] == 0) {
            test_result = {
                status: 0,
                rows: 0
            };
            NLCUTIL_log_classify(log_set, test_set, test_result);

        } else {

            if (hasError == 1) {
                test_result = {
                    status: err_res.status,
                    rows: res_rows[k],
                    nlc: err_res
                };
                NLCUTIL_log_classify(log_set, test_set, test_result);

            } else {
                test_result = {
                    status: nlc_res.status,
                    rows: res_rows[k],
                    nlc: nlc_res
                };
                NLCUTIL_log_classify(log_set, test_set, test_result);
            }
        }
    }
    return
}

