var table = new Table_bs($('#rwlb'));
function cleardata(a) {
    $(a).on("hidden.bs.modal", function () {
        $(this).removeData("bs.modal");
    });

}
function Table_bs($box) {

    var _init_bit = false;
    var _box = $box;

    this.init = function () {

        _box.addClass('table-responsive');

        var html = '';
        html += '<table class="table table-bordered " style="text-align: center; ">';
        html += '<thead style="font-weight: bold;">';
        // html += '<tr>';
        // html += '</tr>';
        html += '</thead>';
        html += '<tbody>';
        // html += '<tr>';
        // html += '</tr>';
        html += '</tbody>';
        html += '</table>';

        _box.append($(html));

        _init_bit = true;
    }
    this.load = function (option) {

        if (!_init_bit) {
            this.init();
        }

        // 表头
        var thead = option.colNames;

        var html_thead = '';
        html_thead += '<tr>';
        for (var i = 0; i < thead.length; i++) {
            html_thead += '<td>' + thead[i] + '</td>';
        }
        html_thead += '</tr>';

        _box.find('thead').append($(html_thead));

        // 表身
        var tbody = option.data;
        var html_tbody = '';

        for (var j1 = 0; j1 < tbody.length; j1++) {
            var buff = tbody[j1];

            html_tbody += '<tr>';
            for (var j2 = 0; j2 < buff.length; j2++) {
                html_tbody += '<td>' + buff[j2] + '</td>';
            }
            html_tbody += '</tr>';
        }
        _box.find('tbody').append($(html_tbody));
    }
    this.reflush = function (option) {
        _box.empty();
        this.init();
        //option.data = [];
        //this.load(option);
    }
}
function del_rw(a) {
    var r = confirm("确认删除任务");
    if (r == true) {
        $.ajax({
            url: "/zb/del_rw",
            data: { 'rwid': a },
            type: "post",
            success: function (data) {
                if (data.IsSuccessful) {
                    alert("删除成功");
                    table.reflush(option);
                    init1();
                }
                else alert(data.ErrMsg);
            }
        });
    }


}


var option = {
    colNames: ['选择', "任务名称", "年份", "周次", "任务状态", "操作", "完整稿", ""],
    data: [

    ],
    url:'',
    param:[]
}

$(function () {
    init1();
});
function init1() {
    $.ajax({
        url: option.url,
        data: option.param,
        type: "post",
        success: function (data) {
            var d = [];
            for (var i = 0; i < data.length; i++) {
                switch (data[i].zt) {
                    case 0:;
                    case 1: {
                        d.push(['<input type="checkbox">', data[i].rwmc, data[i].nf, data[i].zc, "数据筹备中",
                            '<a class="btn btn-primary waves-effect waves-light" href="/data/zb_data?nf=' + data[i].nf + '&zc=' + data[i].zc + '" data-toggle="modal" data-target="#qrsj" onclick="cleardata(\'#qrsj\')">确认数据</a>',

                             '<a class="btn btn-link" onclick="del_rw(' + data[i].rwid + ')">上传定稿</a><a class="btn  btn-link" onclick="del_rw(' + data[i].rwid + ')">下载定稿</a>',
                             '<a class="btn-link" onclick="del_rw(' + data[i].rwid + ')">删除</a>'
                        ]);
                    }; break;
                    case 2: {
                        d.push(['<input type="checkbox">', data[i].rwmc, data[i].nf, data[i].zc, "参数填报中",
                            '<a class="btn btn-primary waves-effect waves-light" href="/zb/add_cs?mbbh=@this.ViewBag.mbbh&rwid=' + data[i].rwid + '" data-toggle="modal" data-target="#tbcs"  onclick="cleardata(\'#tbcs\')">填写参数</a>',

                                 '<a class="btn btn-link" onclick="del_rw(' + data[i].rwid + ')">上传定稿</a><a class="btn  btn-link" onclick="del_rw(' + data[i].rwid + ')">下载定稿</a>',
                                  '<a class="btn-link" onclick="del_rw(' + data[i].rwid + ')">删除</a>'
                        ]);
                    }; break;

                    case 3: {
                        d.push(['<input type="checkbox">', data[i].rwmc, data[i].nf, data[i].zc, "可生成PTT",
                            '<a class="btn btn-primary waves-effect waves-light" href="/zb/sc?mbid=' + data[i].mbid + '&nf=' + data[i].nf + '&zc=' + data[i].zc + '">开始任务</a>',

                             '<a class="btn btn-link" onclick="del_rw(' + data[i].rwid + ')">上传定稿</a><a class="btn  btn-link" onclick="del_rw(' + data[i].rwid + ')">下载定稿</a>',
                             '<a class="btn-link" onclick="del_rw(' + data[i].rwid + ')">删除</a>'
                        ]);
                    }; break;
                    case 4: {
                        d.push(['<input type="checkbox">', data[i].rwmc, data[i].nf, data[i].zc, "可下载",
                            '<a class="btn btn-link" href="/zb/sc?mbid=' + data[i].mbid + '&nf=' + data[i].nf + '&zc=' + data[i].zc + '">重新生成</a><a class="btn btn-link " href="/zb/export?rwid=' + data[i].rwid + '">下载文件</a>',

                            '<a class="btn btn-link" onclick="del_rw(' + data[i].rwid + ')">上传定稿</a>',
                            '<a class="btn-link" onclick="del_rw(' + data[i].rwid + ')">删除</a>'

                        ]);
                    }; break;
                }
            }
            option.data = d;
            table.load(option);
        }
    })

}