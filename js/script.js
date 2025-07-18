/*jslint white: true, browser: true, undef: true, nomen: true, eqeqeq: true, plusplus: false, bitwise: true, regexp: true, strict: true, newcap: true, immed: true, maxerr: 14 */
/*global window: false, REDIPS: true */

/* enable strict mode */
"use strict";

// define init and show methods
var redipsInit,
    getContent;

// redips initialization
redipsInit = function () {
    var num = 0,            // number of successfully placed elements
        rd = REDIPS.drag;   // reference to the REDIPS.drag lib
    rd.init();              // initialization
    rd.hover.colorTd = '#9BB3DA';   // set hover color
    // on each drop refresh content
    rd.event.dropped = function () {
        // ドロップ先 td の id を求める
        var tdId = rd.td.target.id;
        // ドロップしたdivのIDを求める
        var i, j, cn, gcn, contentId, contentValue;
        for (i = 0; i < rd.td.target.childNodes.length; i++) {
            cn = rd.td.target.childNodes[i];
            if (cn.nodeName === 'DIV') {
                contentId = cn.id;
                for (j = 0; j < cn.childNodes.length; j++) {
                    gcn = cn.childNodes[j];
                    if (gcn.nodeName === 'INPUT') {
                        // ドロップした場所に関連した名称にドロップした項目名を変更
                        contentValue = gcn.value;
                        // value値書き換え(オペレータ社員コード_tdId)
                        gcn.value = contentValue.split("_")[0] + "_" + tdId;
                    }
                }
            }
        }
    };
};

// get content (DIV elements in TD)
getContent = function (id) {
    var td = document.getElementById(id),
        content = '',
        cn, i;
    // TD can contain many DIV elements
    for (i = 0; i < td.childNodes.length; i++) {
        // set reference to the child node
        cn = td.childNodes[i];
        // childNode should be DIV with containing "drag" class name
        if (cn.nodeName === 'DIV' && cn.className.indexOf('drag') > -1) { // and yes, it should be uppercase
            // append DIV id to the result string
            content += cn.id + '_';
        }
    }
    // cut last '_' from string
    content = content.substring(0, content.length - 1);
    // return result
    return content;
};

// add onload event listener
if (window.addEventListener) {
    window.addEventListener('load', redipsInit, false);
}
else if (window.attachEvent) {
    window.attachEvent('onload', redipsInit);
}

// 集計
function setTotalUp(){
    
}
