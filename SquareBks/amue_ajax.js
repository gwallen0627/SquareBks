function ajx(m, d, cb) {
        var loc = location.pathname.substring(1);
    if (loc.indexOf(".") == -1) { loc += ".aspx" }
    $.ajax(
        { type: "POST", url: loc + "/" + m, data: d, contentType: "application/json; charset=utf-8", dataType: "json", success: cb }
    );
}
function dodata(form) { var o = {}; var a = $("#" + form).serializeArray(); $.each(a, function () { if (o[this.name]) { if (!o[this.name].push) { o[this.name] = [o[this.name]]; } o[this.name].push(this.value || ''); } else { o[this.name] = this.value || ''; } }); return o; };
