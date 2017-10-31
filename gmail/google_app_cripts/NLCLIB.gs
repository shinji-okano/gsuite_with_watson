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
// グローバル変数
var CLFNAME_PREFIX = "CLF";
var CLF_SEP = "#__#";

var NLCUTIL_NB_CONF = 13;
var NB_NOTIF_CONF = 8;

var NOTIF_OPT = {
    ON: "On",
    OFF: "Off"
};
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// クレデンシャルの取得
function NLCUTIL_load_creds() {

    var scriptProps = PropertiesService.getScriptProperties();

    var creds = {};

    creds["url"] = scriptProps.getProperty('CREDS_URL');
    creds["username"] = scriptProps.getProperty('CREDS_USERNAME');
    creds["password"] = scriptProps.getProperty('CREDS_PASSWORD');

    return creds;
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// 設定情報の取得
function NLCUTIL_load_config(config_set, ss_id, ws_name, start_col, start_row) {

    var ss = SpreadsheetApp.openById(config_set.ss_id);

    var sheet = ss.getSheetByName(config_set.ws_name);

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var records = sheet.getRange(config_set.st_start_row, config_set.st_start_col, lastRow - config_set.st_start_row + 1, NLCUTIL_NB_CONF).getValues();

    var i = 0;
    var sheet_conf = {
        ws_name: records[CONF_INDEX.ws_name][i],
        start_row: records[CONF_INDEX.start_row][i],
        start_col: records[CONF_INDEX.start_col][i],
        text_col: records[CONF_INDEX.text_col][i],
        intent_col: [records[CONF_INDEX.intent1_col][i], records[CONF_INDEX.intent2_col][i], records[CONF_INDEX.intent3_col][i]],
        result_col: [records[CONF_INDEX.result1_col][i], records[CONF_INDEX.result2_col][i], records[CONF_INDEX.result3_col][i]],
        restime_col: [records[CONF_INDEX.restime1_col][i], records[CONF_INDEX.restime2_col][i], records[CONF_INDEX.restime3_col][i]],
        log_ws: records[CONF_INDEX.log_ws][i]
    };

    var notif_conf = {};
    notif_conf["option"] = records[CONF_INDEX.notif_opt][0];
    notif_conf["ws_name"] = records[CONF_INDEX.notif_ws][0];

    return {
        sheet_conf: sheet_conf,
        notif_conf: notif_conf,
    };
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
//通知条件設定の取得
function NLCUTI_load_notif_rules(config_set) {

    var ss = SpreadsheetApp.openById(config_set.ss_id);

    var sheet = ss.getSheetByName(config_set.ws_name);

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow == 0 || lastCol == 0) return [];


    var records = sheet.getRange(config_set.start_row, config_set.start_col, lastRow - config_set.start_row + 1, NB_NOTIF_CONF).getValues();

    var rules = [];

    for (var i = 0; i < records.length; i++) {

        rules.push({
            res_int: [String(records[i][0]), String(records[i][1]), String(records[i][2])],
            mail: {
                to: records[i][3],
                cc: records[i][4],
                bcc: records[i][5],
                subject: records[i][6],
                body: records[i][7]
            }
        });

    }
    return rules;
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// メールの送信
function NLCUTIL_send_mail(mail_set) {

    Logger.log("### NLCUTIL_notify");

    GmailApp.sendEmail(
        mail_set.to,
        mail_set.subject,
        mail_set.body, {
            from: mail_set.from,
            name: mail_set.sender
        }
    );
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
function NLCUTIL_check_notify(notif_set, record, upd_flg) {

    Logger.log("### NLCUTIL_check_notify");

    for (var i = 0; i < notif_set.rules.length; i++) {

        var chk_cnt = 0;
        var upd_chk = 0;
        for (var j = 0; j < NB_CLFS; j++) {
      var HOGE1 = notif_set.rules[i].res_int[j]
      var HOGE2 = record[notif_set.result_cols[j] - 1] 
      var HOGE3 = notif_set.rules[i].res_int[j]

// WILD
            if ("" == notif_set.rules[i].res_int[j]) {
                chk_cnt += 1;
            } else {
                if (upd_flg[j] == 1) {
                    upd_chk = 1;
                }
                if (record[notif_set.result_cols[j] - 1] == notif_set.rules[i].res_int[j]) {
                    chk_cnt += 1;
                }
            }
        }

        if (chk_cnt == NB_CLFS && 1 == upd_chk) {

            var res;
            res = NLCUTIL_expand_tags(notif_set.rules[i].mail.body, record);
            var body = res.text;

            res = NLCUTIL_expand_tags(notif_set.rules[i].mail.subject, record);
            var subject = res.text;

            var mail_set = {
                from: notif_set.from,
                sender: notif_set.sender,
                to: notif_set.rules[i].mail.to,
                subject: subject,
                body: body
            };

            NLCUTIL_send_mail(mail_set);
        }
    }
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// 通知条件のチェック
function NLCUTIL_check_notif(notif_set, data_set) {

    Logger.log("### NLCUTIL_check_notif");

    var ss = SpreadsheetApp.openById(data_set.ss_id);

    var sheet = ss.getSheetByName(data_set.ws_name);

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries = sheet.getRange(data_set.start_row, data_set.start_col, lastRow - data_set.start_row + 1, lastCol - data_set.start_col + 1).getValues();

    for (var i = 0; i < notif_set.rules.length; i++) {

        var isMatch = entries.some(function(element, index, array) {

            return (
                element[data_set.result_col[0] - data_set.start_col] == notif_set.rules[i].res_int[0] &&
                element[data_set.result_col[1] - data_set.start_col] == notif_set.rules[i].res_int[1] &&
                element[data_set.result_col[2] - data_set.start_col] == notif_set.rules[i].res_int[2]
            );
        });

        if (true == isMatch) {

            var mail_set = {
                from: notif_set.from,
                sender: notif_set.sender,
                to: notif_set.rules[i].mail.to,
                subject: notif_set.rules[i].mail.subject,
                body: notif_set.rules[i].mail.body
            };

            NLCUTIL_send_mail(mail_set);
        }
    }
}
// ----------------------------------------------------------------------------




// ----------------------------------------------------------------------------
// 分類処理
function NLCUTIL_classify(test_set, creds_username, creds_password, override) {

    Logger.log("### NLCUTIL_classify");

    var log_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.log_ws,
        start_col: CONFIG_SET.log_start_col,
        start_row: CONFIG_SET.log_start_row,
    };

    var clf = NLCUTIL_select_clf(test_set.clf_name, creds_username, creds_password);
    if ("Training" == clf.status) {
        NLCUTIL_log_classify(log_set, test_set, {
            status: "900",
            description: "トレーニング中",
            clf_id: clf.clf_id
        });
    } else if ("Nothing" == clf.status) {
        NLCUTIL_log_classify(log_set, test_set, {
            status: 900,
            description: "分類器なし",
            clf_id: ""
        });
    } else if ("Available" != clf.status) {
        NLCUTIL_log_classify(log_set, test_set, {
            status: "900",
            description: "分類器なし",
            clf_id: ""
        });
    }

    var ss = SpreadsheetApp.openById(test_set.ss_id);

    var sheet = ss.getSheetByName(test_set.ws_name);

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries = sheet.getRange(test_set.start_row, test_set.start_col, lastRow - test_set.start_row + 1, lastCol - test_set.start_col + 1).getValues();

    var hasError = 0;
    var results = [];
    var inres = [];
    var restimes = [];
    var row_cnt = 0;
    var updates = 0;
    var upd_flg = [0, 0, 0];
    for (var i = 0; i < entries.length; i++) {

        if (lastCol < test_set.result_col) {
            result_text = "";
        } else {
            result_text = entries[cnt][test_set.result_col - test_set.start_col];
        }

        if (result_text != "" && override != true) continue;

        var result_text;
        test_text = entries[i][test_set.text_col - 1];
        test_text = NLCUTIL_norm_text(test_text);

        if (0 == test_text.length) continue;

        if (test_text.length > 1024) {
            test_text = test_text.substring(0, 1024);
        }

        var res = NLCAPI_post_classify(creds_username, creds_password, clf_id, test_text);
        if (nlc_res.status != 200) {
            hasError = 1;
            err_res = nlc_res;
        } else {
            var r = res.body.top_class;
            sheet.getRange(test_set.start_row + cnt, test_set.result_col, 1, 1).setValue(r);
            var t = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
            sheet.getRange(test_set.start_row + cnt, test_set.restime_col, 1, 1).setValue(t);
            row_cnt += 1;
            updates = 1;
            upd_flg[test_set.clf_no - 1] = 1;

            if (NOTIF_OPT.ON == test_set.notif_opt) {
                var record = sheet.getRange(test_set.start_row + cnt, 1, 1, lastCol).getValues();
                NLCUTIL_check_notify(test_set.notif_set, record[0], upd_flg);
            }
        }
    }

    var clfname = CLFNAME_PREFIX + String(test_set.clf_no);
    var test_result;
    if (0 == row_cnt) {
        test_result = {
            status: 0,
            rows: 0
        };
        NLCUTIL_log_classify(log_set, test_set, test_result);
    } else {
        if (1 == hasError) {
            test_result = {
                status: err_res.status,
                rows: row_cnt,
                nlc: err_res
            };
            NLCUTIL_log_classify(log_set, test_set, test_result);
        } else {
            test_result = {
                status: nlc_res.status,
                rows: row_cnt,
                nlc: nlc_res
            };
            NLCUTIL_log_classify(log_set, test_set, test_result);
        }
    }
    return;
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// イベント実行用分類処理
function NLCUTIL_classify_all() {

    Logger.log("### NLCUTIL_classify_all");

    var conf = NLCUTIL_load_config(CONFIG_SET);

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
        text_col: conf.sheet_conf.text_col,
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
        if ("Training" == clf.status) {
            NLCUTIL_log_classify(log_set, test_set, {
                status: 900,
                description: "トレーニング中",
                clf_id: clf.clf_id
            });
        } else if ("Nothing" == clf.status) {
            NLCUTIL_log_classify(log_set, test_set, {
                status: 900,
                description: "分類器なし",
                clf_id: ""
            });
        } else if ("Available" != clf.status) {
            NLCUTIL_log_classify(log_set, test_set, {
                status: 800,
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

        var test_text = entries[cnt][conf.sheet_conf.text_col - conf.sheet_conf.start_col];
        test_text = NLCUTIL_norm_text(String(test_text)).trim();
        if (0 == test_text.length) continue;

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
        if (0 == res_rows[k]) {
            test_result = {
                status: 0,
                rows: 0
            };
            NLCUTIL_log_classify(log_set, test_set, test_result);

        } else {

            if (1 == hasError) {
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
    return;
}
// ----------------------------------------------------------------------------




// ----------------------------------------------------------------------------
// 学習処理
function NLCUTIL_train(train_set, creds_username, creds_password) {

    var clfs = NLCAPI_get_classifiers(creds_username, creds_password);
    if (clfs.status != 200 ) {
        return {
            status: clfs.status,
            description: clfs.error,
        };
    }

    var ss = SpreadsheetApp.openById(train_set.ss_id);

    var sheet = ss.getSheetByName(train_set.ws_name);

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries = sheet.getRange(train_set.start_row, 1, lastRow - train_set.start_row + 1, lastCol).getValues();

    var row_cnt = 0;
    var csvString = "";
    for (var i = 0; i < entries.length; i++) {

        var train_text = String(entries[i][train_set.text_col - 1]);
        train_text = NLCUTIL_norm_text(train_text);

        if (train_text.length > 1024) {
            train_text = train_text.substring(0, 1024);
        }

        var class_name = NLCUTIL_norm_text(String(entries[i][train_set.class_col - 1]));

        if (train_text.length > 0 && class_name.length > 0) {
            csvString = csvString + '"' + train_text + '","' + class_name + '"' + "\r\n";
            row_cnt += 1;
        }

    }


    var clf_info = NLCUTIL_clf_vers(clfs.body.classifiers, train_set.clf_name);

    var new_version = (clf_info.max_ver + 1);
    var clf_name = train_set.clf_name + "#__#" + new_version;

    var res = NLCAPI_post_classifiers(creds_username, creds_password, csvString, clf_name, 'ja');

    if (clf_info.count >= 2 && 200 == res.status) {

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
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// 特定分類器で学習して結果をログに出力する

function NLCUTIL_train_set(clf_no) {

    Logger.log("### NLCUTIL_train_set", clf_no);

    var conf = NLCUTIL_load_config(CONFIG_SET);

    var train_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_row: conf.sheet_conf.start_row,
        start_col: conf.sheet_conf.start_col,
        end_row: -1,
        text_col: conf.sheet_conf.text_col,
        class_col: conf.sheet_conf.intent_col[clf_no - 1],
        clf_no: clf_no,
        clf_name: CLFNAME_PREFIX + clf_no,
    };

    var train_result = NLCUTIL_train(train_set, CREDS.username, CREDS.password);

    var log_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.log_ws,
        start_col: CONFIG_SET.log_start_col,
        start_row: CONFIG_SET.log_start_row
    };

    NLCUTIL_log_train(log_set, train_set, train_result);
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// イベント実行用学習処理
function NLCUTIL_train_all() {

    for (var i = 1; i <= NB_CLFS; i++) {
        NLCUTIL_train_set(i);
    }

}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// 分類器の利用可能な最新バージョンを取得する
function NLCUTIL_select_clf(clf_name, creds_username, creds_password) {

    var clfs = NLCAPI_get_classifiers(creds_username, creds_password);
    if ( clfs.status != 200 ) {
        return {
            clf_id: "",
            status: clfs.status,
            nlc: clfs.body,
        };
    }

    var clf_info = NLCUTIL_clf_vers(clfs.body.classifiers, clf_name);

    if (0 == clf_info.count) {
        return {
            clf_id: "",
            status: "Nothing"
        };
    }

    var clf;
    var res;
    for (var i = clf_info.max_ver; i >= clf_info.min_ver; i--) {

        clf = clf_info.clfs[i];

        res = NLCAPI_get_classifier(creds_username, creds_password, clf.classifier_id);

        if ("Available" == res.body.status) {
            return {
                clf_id: clf.classifier_id,
                status: res.body.status
            };
        }
    }

    return {
        clf_id: clf.classifier_id,
        status: res.body.status
    };
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// 分類器のバージョン一覧を取得
function NLCUTIL_clf_vers(clf_list, target_name) {

    var clfs = [];
    var max_ver = 0;
    var min_ver = 99999999;
    var count = 0;

    for (var i = 0; i < clf_list.length; i++) {

        var base = clf_list[i].name.split(CLF_SEP);
        if (null == base[1]) continue;
        if (target_name == base[0]) {

            clfs[parseInt(base[1])] = clf_list[i];
            count += 1;

            if (parseInt(base[1]) > max_ver) {
                max_ver = parseInt(base[1]);
            }
            if (parseInt(base[1]) < min_ver) {
                min_ver = parseInt(base[1]);
            }
        }
    }

    var result = {
        count: count,
        min_ver: min_ver,
        max_ver: max_ver,
        clfs: clfs
    };
    return result;
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
function NLCUTIL_log_train(log_set, train_set, res) {

    var ss = SpreadsheetApp.openById(log_set.ss_id);

    var sheet = ss.getSheetByName(log_set.ws_name);
    if (null == sheet) {
        sheet = ss.insertSheet(log_set.ws_name);
    }

    var lastRow = sheet.getLastRow();
    lastRow = lastRow + 1;
    if (lastRow < log_set.start_row) {
        lastRow = log_set.start_row;
    }

    var lastCol = sheet.getLastColumn();

    var record = [];
    if (200 == res.status) {
        record = ["学習",
            Utilities.formatDate(new Date(res.nlc.from), "JST", "yyyy/MM/dd HH:mm:ss"),
            res.nlc.body.status,
            train_set.clf_no,
            train_set.ws_name,
            res.rows,
            train_set.text_col,
            train_set.class_col,
            res.nlc.body.classifier_id,
            res.status,
            res.nlc.body.status_description,
            res.nlc.body.created,
            res.version
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length).setValues([record]);
    } else {
        if (res["nlc"]) {

            record = ["学習", 　
                Utilities.formatDate(new Date(res.nlc.from), "JST", "yyyy/MM/dd HH:mm:ss"),
                res.nlc.body.error,
                train_set.clf_no,
                train_set.ws_name,
                res.rows,
                train_set.text_col,
                train_set.class_col,
                "",
                res.status,
                res.nlc.body.description
            ];
        } else {
            record = ["学習", 　
                Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
                res.description,
                train_set.clf_no,
                train_set.ws_name,
                "N/A",
                train_set.text_col,
                train_set.class_col,
                "",
                res.status,
            ];
        }
        sheet.getRange(lastRow, log_set.start_col, 1, record.length).setValues([record]).setFontColor("red");
    }

}
// ----------------------------------------------------------------------------



// ----------------------------------------------------------------------------
function NLCUTIL_log_classify(log_set, test_set, test_result) {

    Logger.log("### NLCUTIL_log_classify");

    var ss = SpreadsheetApp.openById(log_set.ss_id);

    var sheet = ss.getSheetByName(log_set.ws_name);
    if (null == sheet) {
        sheet = ss.insertSheet(log_set.ws_name);
    }

    var lastRow = sheet.getLastRow();
    lastRow = lastRow + 1;
    if (lastRow < log_set.start_row) {
        lastRow = log_set.start_row;
    }

    var lastCol = sheet.getLastColumn();

    var record = [];
    if (200 == test_result.status) {
        record = ["分類",
            Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
            "成功",
            test_set.clf_no,
            test_set.ws_name,
            test_result.rows,
            test_set.text_col,
            test_set.result_col,
            test_result.nlc.body.classifier_id,
            test_result.status
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length).setValues([record]);
    } else if (0 == test_result.status) {
        record = [
            "分類",
            Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
            "対象なし",
            test_set.clf_no,
            test_set.ws_name,
            test_result.rows,
            test_set.text_col,
            test_set.result_col
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length).setValues([record]);
    } else if (900 == test_result.status) {
        record = [
            "分類",
            Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
            test_result.description,
            test_set.clf_no,
            test_set.ws_name,
            "N/A",
            test_set.text_col,
            test_set.result_col,
            test_result.clf_id,
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length).setValues([record]);
    } else {
        if (test_result["nlc"]) {
            record = ["分類",
                Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
                test_result.description,
                test_set.clf_no,
                test_set.ws_name,
                "N/A",
                test_set.text_col,
                test_set.result_col,
            ];
        } else {
            record = ["分類",
                Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
                test_result.description,
                test_set.clf_no,
                test_set.ws_name,
                "N/A",
                test_set.text_col,
                test_set.result_col,
            ];
        }
        sheet.getRange(lastRow, log_set.start_col, 1, record.length).setValues([record]).setFontColor("red");
    }

}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
function NLCUTIL_exec_check_clfs() {

    var clf_set = {
        ss_id: CONFIG_SET.ss_id,
        ws_name: CONFIG_SET.ws_name,
        start_col: CONFIG_SET.clfs_start_col,
        start_row: CONFIG_SET.clfs_start_row,
    };

    NLCUTIL_check_classifiers(clf_set, CREDS);
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
function NLCUTIL_check_classifiers(clf_set, creds) {

    Logger.log("### NLCUTIL_check_classifiers");

    var ss = SpreadsheetApp.openById(clf_set.ss_id);

    var sheet = ss.getSheetByName(clf_set.ws_name);

    sheet.getRange(clf_set.start_row, clf_set.start_col, NB_CLFS, 2).clear();

    var clfs = NLCAPI_get_classifiers(creds.username, creds.password);

    CLFS_ALL_STATUS = 0;
    for (var cnt = 1; cnt <= NB_CLFS; cnt++) {

        if (clfs.status != 200) {
            sheet.getRange(clf_set.start_row + cnt - 1, clf_set.start_col, 1, 2).setValues([
                ["ERROR", clfs.error]
            ]);
        } else {

            var clf_name = CLFNAME_PREFIX + String(cnt);

            var clf_info = NLCUTIL_clf_vers(clfs.body.classifiers, clf_name);

            if (0 == clf_info.count) {
                sheet.getRange(clf_set.start_row + cnt - 1, clf_set.start_col, 1, 2).setValues([
                    ["", ""]
                ]);
                CLFS_ALL_STATUS += 1;
            } else {
                var clf = clf_info.clfs[clf_info.max_ver];

                var res = NLCAPI_get_classifier(creds.username, creds.password, clf.classifier_id);
                if ("Available" == res.body.status) {
                    CLFS_ALL_STATUS += 1;
                }

                sheet.getRange(clf_set.start_row + cnt - 1, clf_set.start_col, 1, 2).setValues([
                    [clf.classifier_id, res.body.status]
                ]);
            }
        }
    }

    // 全てAvailableでタイマー解除
    if (NB_CLFS == CLFS_ALL_STATUS) {
        NLCUTIL_del_trigger("NLCUTIL_exec_check_clfs");
    }
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// タイマーをセットする
function NLCUTIL_set_trigger(func_name, int_min) {

    var triggers = ScriptApp.getProjectTriggers();

    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() == func_name) {
            return;
        }
    }

    var triggerDay = new Date();
    ScriptApp.newTrigger(func_name)
        .timeBased()
        .everyMinutes(int_min)
        .create();
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// タイマーを取り消す
function NLCUTIL_del_trigger(func_name) {

    Logger.log("### NLCUTIL_del_trigger");

    var triggers = ScriptApp.getProjectTriggers();

    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() == func_name) {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// 埋め込みタグを展開する
function NLCUTIL_expand_tags(target, record) {

    var xbody = target;
    var buf = "";

    for (var idx = 0; idx < record.length; idx++) {
        buf = xbody.replace(new RegExp('\\[\\[#' + String(idx + 1) + '\\]\\]', 'g'), record[idx]);
        xbody = buf;
    }

    var errtag = xbody.match(new RegExp('\[\[.+\]\]', 'g'));
    if (errtag !== null) {
        return {
            'code': 'ERROR',
            'errtag': errtag
        };
    }

    return {
        'code': 'OK',
        'text': xbody
    };
}
// ----------------------------------------------------------------------------





// ----------------------------------------------------------------------------
// テキストをノーマライズする
function NLCUTIL_norm_text(target) {
    var res = target
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\"\'")
        .replace(/"/g, '\"\"')
        .replace(/\//g, '\\/')
        .replace(/</g, '\\x3c')
        .replace(/>/g, '\\x3e')
        .replace(/(0x0D)/g, '')
        .replace(/(0x0A)/g, '')
        .replace(/\t/g, ' ')
        .replace(/\r?\n/g, "")
        .replace(/\r/g, "")
        .replace(/\n/g, "");

    res = res.trim();

    return res;
}
// ----------------------------------------------------------------------------





// Watson Natural Language Classifier API wrapper
// ----------------------------------------------------
// 定数
var URI_DOMAIN = "https://gateway.watsonplatform.net";
var URI_BASE = "/natural-language-classifier/api";
var URI_APIVERSION = "v1";
var EOL = '\r\n';
// DataCollection
//X-Watson-Learning-Opt-Out true
// ----------------------------------------------------


// ----------------------------------------------------
// Classifier一覧取得
// GET /v1/classifiers
// 概要: NLCサービスのClassifier一覧を取得する
// パラメータ
// username: NLCサービス資格情報のusername
// password: NLCサービス資格情報のpassword
// 戻り値: APIのレスポンスボディ
function NLCAPI_get_classifiers(p_username, p_password) {

    // URIビルド
    var resource = "classifiers";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource;
    // リクエストオプション
    var options = {
        "headers": {
            "Authorization": " Basic " + Utilities.base64Encode(p_username + ":" + p_password)
        },
        "method": "get",
        "contentType": "application/json",
        "muteHttpExceptions": true
    };

    var fromTime = new Date();

    // HTTPリクエスト
    var response = UrlFetchApp.fetch(uri, options);

    // HTTPレスポンスをJSONに変換
    var result = JSON.parse(response);
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText()
    var result;
    result = {
        status: responseCode,
        body: JSON.parse(responseBody),
        from: fromTime.getTime(),
        to: toTime.getTime(),
        delta: delta
    };
    // 戻り値を返す
    return result;
}
// ----------------------------------------------------



// ----------------------------------------------------
// Classifier生成
// POST /v1/classifiers
// 概要: 資格情報に該当するNLCサービスにClassifierを新規作成する
// パラメータ
// username: NLCサービス資格情報のusername
// password: NLCサービス資格情報のpassword
// p_training_data: トレーニングデータ(CSV)
// p_classname: クラス名(training_metadata)
// p_langcode: 言語コード(オプション)
// 戻り値: APIのレスポンスボディ
function NLCAPI_post_classifiers(p_username, p_password, p_training_data, p_classname, p_langcode) {

    // 日本語のデフォルト(ja)設定
    p_langcode = p_langcode || "ja";

    // URIビルド
    var resource = "classifiers";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource;

    // メタデータ
    var training_metadata = {
        "language": p_langcode,
        "name": p_classname
    };

    // バウンダリデータの生成
    var boundary = NLCAPI_createBoundary();

    // バイナリデータの生成
    var postbody = Utilities.newBlob('--' + boundary + EOL +
        'Content-Disposition: form-data; name="training_data"; filename="training.csv"' + EOL +
        'Content-Type: application/octet-stream' + EOL +
        'Content-Transfer-Encoding: binary' + EOL + EOL).getBytes();
    postbody = postbody.concat(Utilities.newBlob(p_training_data).getBytes());
    postbody = postbody.concat(Utilities.newBlob('--' + boundary + EOL +
        'Content-Disposition: form-data; name="training_metadata"' + EOL + EOL).getBytes());
    postbody = postbody.concat(Utilities.newBlob(JSON.stringify(training_metadata)).getBytes());
    postbody = postbody.concat(Utilities.newBlob(EOL + '--' + boundary + '--' + EOL).getBytes());
    //Logger.log( String.fromCharCode.apply( null, postbody ) );

    // リクエストオプション
    var options = {
        "headers": {
            "Authorization": " Basic " + Utilities.base64Encode(p_username + ":" + p_password)
        },
        "method": "post",
        "contentType": 'multipart/form-data; boundary=' + boundary,
        "payload": postbody,
        "muteHttpExceptions": true
    };

    var fromTime = new Date();

    var response;
    try {
        response = UrlFetchApp.fetch(uri, options);
    } catch (e) {
        Logger.log(e)
        return {
            status: 999,
            description: e
        }
    }

    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText()
    var result;

    if (responseCode != 200) {
        result = {
            status: responseCode,
            body: JSON.parse(responseBody),
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        return result;
    } else {
        result = {
            status: responseCode,
            body: JSON.parse(responseBody),
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        return result;
    }
}
// ----------------------------------------------------



// ----------------------------------------------------
// クラス分類
// POST /v1/classifiers/{classifier_id}/classify
// 概要: 文章を対象の分類器で分類する
// パラメータ
// username: NLCサービス資格情報のusername
// password: NLCサービス資格情報のpassword
// classid: 分類器のクラスID
// p_phrase: 分類する文章
// 戻り値: APIのレスポンスボディ
function NLCAPI_post_classify(p_username, p_password, p_classid, p_phrase) {

    // URIビルド
    var resource = "classifiers";
    var verb = "classify";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource + "/" + p_classid + "/" + verb;

    // ポストデータ
    var postbody = JSON.stringify({
        "text": p_phrase
    });

    // リクエストオプション
    var options = {
        "headers": {
            "Authorization": " Basic " + Utilities.base64Encode(p_username + ":" + p_password)
        },
        "method": "post",
        "contentType": "application/json",
        "payload": postbody,
        "muteHttpExceptions": true
    };

    var fromTime = new Date();
    var response = UrlFetchApp.fetch(uri, options);
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText()
    var result;

    if (responseCode != 200) {
        result = {
            status: responseCode,
            body: responseBody,
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        return result;
    } else {
        result = {
            status: responseCode,
            body: JSON.parse(responseBody),
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        result = {
            status: responseCode,
            body: JSON.parse(responseBody),
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        // 戻り値を返す
        return result;
    }
}
// ----------------------------------------------------


// ----------------------------------------------------
// クラス分類
// GET /v1/classifiers/{classifier_id}/classify
// 概要: 文章を対象のClassifierで分類する
// パラメータ
// p_username: NLCサービス資格情報のusername
// p_password: NLCサービス資格情報のpassword
// p_classid: Classifier_id
// p_phrase: 分類する文章
// 戻り値: APIのレスポンスボディ
function NLCAPI_get_classify(p_username, p_password, p_classid, p_phrase) {

    // URIビルド
    var resource = "classifiers";
    var verb = "classify";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource + "/" + p_classid + "/" + verb;
    uri = uri + "?text=" + p_phrase;

    // リクエストオプション
    var options = {
        "headers": {
            "Authorization": " Basic " + Utilities.base64Encode(p_username + ":" + p_password)
        },
        "method": "get",
        "contentType": "application/json",
        "muteHttpExceptions": true
    };

    var fromTime = new Date();
    var response = UrlFetchApp.fetch(uri, options);
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText()
    var result;

    if (responseCode != 200) {
        result = {
            status: responseCode,
            body: responseBody,
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        return result;
    } else {
        result = {
            status: responseCode,
            body: JSON.parse(responseBody),
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        result = {
            status: responseCode,
            body: JSON.parse(responseBody),
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        // 戻り値を返す
        return result;
    }
}
// ----------------------------------------------------



// ----------------------------------------------------
// Classifier削除
// DELETE /v1/classifiers/{classifier_id}
// 概要: 対象のClassifierを削除する
// パラメータ
// username: NLCサービス資格情報のusername
// password: NLCサービス資格情報のpassword
// classid: 削除対象のClassifier_id
// 戻り値: APIのレスポンスボディ
function NLCAPI_delete_classifier(p_username, p_password, p_classid) {

    // URIビルド
    var resource = "classifiers";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource + "/" + p_classid;
    // リクエストオプション
    var options = {
        "headers": {
            "Authorization": " Basic " + Utilities.base64Encode(p_username + ":" + p_password)
        },
        "method": "delete",
        "contentType": "application/json",
        "muteHttpExceptions": true
    };

    // HTTPリクエスト
    var fromTime = new Date();
    var response = UrlFetchApp.fetch(uri, options);
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText()
    var result;

    if (responseCode != 200) {
        result = {
            status: responseCode,
            body: responseBody,
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        return result;
    } else {
        result = {
            status: responseCode,
            body: JSON.parse(responseBody),
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        result = {
            status: responseCode,
            body: JSON.parse(responseBody),
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        // 戻り値を返す
        return result;
    }
}
// ----------------------------------------------------



// ----------------------------------------------------
// Classifier情報取得
// GET /v1/classifiers/{classifier_id}
// 概要: 対象のClassifierについてステータスなどの情報を取得する
// パラメータ
// username: NLCサービス資格情報のusername
// password: NLCサービス資格情報のpassword
// classid: 取得対象のClassifier_id
// 戻り値: APIのレスポンスボディ
function NLCAPI_get_classifier(p_username, p_password, p_classid) {

    // URIビルド
    var resource = "classifiers";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource + "/" + p_classid;
    // リクエストオプション
    var options = {
        "headers": {
            "Authorization": " Basic " + Utilities.base64Encode(p_username + ":" + p_password)
        },
        "method": "get",
        "contentType": "application/json",
        "muteHttpExceptions": true
    };

    // HTTPリクエスト
    var fromTime = new Date();
    var response = UrlFetchApp.fetch(uri, options);
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText()
    var result;

    if (responseCode != 200) {
        result = {
            status: responseCode,
            body: responseBody,
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        return result;
    } else {
        result = {
            status: responseCode,
            body: JSON.parse(responseBody),
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        result = {
            status: responseCode,
            body: JSON.parse(responseBody),
            from: fromTime.getTime(),
            to: toTime.getTime(),
            delta: delta
        };
        // 戻り値を返す
        return result;
    }

}
// ----------------------------------------------------



// ----------------------------------------------------
// multipartデータ送信のバウンダリ生成
function NLCAPI_createBoundary() {

    var multipartChars = "-_1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
    var length = 30 + Math.floor(Math.random() * 10);
    var boundary = "---------------------------";
    for (var i = 0; i < length; i++) {
        boundary += multipartChars.charAt(Math.floor(Math.random() * multipartChars.length));
    }
    return boundary;
}
// ----------------------------------------------------

