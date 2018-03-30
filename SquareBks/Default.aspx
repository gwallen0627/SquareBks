<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="SquareBks._Default" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>SquareBooks</title>
    <script type="text/javascript" src="jquery-1.11.0.min.js"></script>
    <script type="text/javascript" src="amue_ajax.js"></script>
    <script type="text/javascript">
        function QBI(sender, args) {
            $(window).load(function () { pset() });
            var f = args.get_fileName();
            $("#ldr").show();
            PageMethods.ProcessQBFile(f, function (r) {
                var ra = r.split("|");
                $("#QBItemsSel").html(ra[0]);
                if (ra[1]) { $("#AddDiv").show(); $("#NewQBItemsSel").html(ra[1]) };
                $("#ldr").fadeOut();
            });
        }
        function AddQBIs() {
            var s = "";
            $("#NewQBItemsSel option").each(function (r) { s += $(this).text() + "|" });
            PageMethods.AddQBItems(s, function (r) { $("#QBItemsSel").html(r); alert("Items Added"); $("#AddDiv").hide() });
        }
        function SQI(sender, args) {
            $("#ldr").show();
            var f = args.get_fileName();
            PageMethods.ProcessSQFile(f, function (r) {
                $("#ldr").fadeOut();
                $("#SQCell").empty();
                if (r) { $("#SQCell").append(r) } else { if (confirm("Produce Output File for '" + f + "'?")) { output(f) } }
            })
        }
        function output(f) { $("#ldr").show(); PageMethods.DoOutput(f, function (r) { var ra = r.split("|"); $("#ldr").fadeOut(); $("#SQCell").append(ra[0]); window.location.assign(encodeURI(ra[1])) }) }
        function SaveSQItems(f) {
            var missing = false;
            $('#SQTable select').each(function () { if ($(this).val() == "") { missing = true } });
            missing = false;
            if (missing) { alert("Some items are missing assignments") } else {
                var args = "";
                $('#SQTable tr').each(function () {
                    var name = $(this).find("td").eq(0).html();
                    var pp = $(this).find("td").eq(1).html();
                    var assid = $(this).find("select").eq(0).val();
                    args += name + "|" + pp + "|" + assid + "////";
                });
                PageMethods.SubmitSQItems(args, function (r) { if (r) { alert(r) } else { $("#SQCell").empty(); if (confirm("Items Assigned Successfully!  Produce Output File for '" + f + "'?")) { output(f) } } })
            }
        }
        function bf_proc(sender, args) {
            $("#ldr").show();
            PageMethods.bf_proc(args.get_fileName(), function (r) {
                $("#ldr").fadeOut();
                $("#bank_cell").html(r);

            });
        }
    </script>

    <style type="text/css">
        td
        {
            text-align: left;
        }

        th
        {
            color: White;
            background-color: Aqua;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <table style="left: 50%; top: 50%; transform: translate(-50%, -50%);-webkit-transform : translate(-50%, -50%); position: fixed;display:none">
            <tr>
                <td>Please wait...<br />
                    <img id="ldr" src="loader.gif" style="width: 30px; height: 30px; display: block; margin: auto" />
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 50%; vertical-align: top; padding: 20px;">
                    <table style="width: 100%">
                        <tr>
                            <td>
                                <span style="font-size: xx-large">Upload QB File</span>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <ajaxToolkit:AjaxFileUpload ID="AjaxFileUpload1" runat="server" OnUploadComplete="AjaxFileUpload1_UploadComplete"
                                    OnClientUploadComplete="QBI" MaximumNumberOfFiles="1" AllowedFileTypes="xls,xlsx" />
                            </td>
                        </tr>
                        <tr>
                            <td>QB Items
                                <select id="QBItemsSel">
                                </select>

                            </td>
                        </tr>
                        <tr>
                            <td>
                                <div id="AddDiv" style="display: none">
                                    <table>
                                        <tr>
                                            <td>New QB Item(s):
                                                <select id="NewQBItemsSel">
                                                </select>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <input type="button" value="Add All" onclick="AddQBIs();" />
                                            </td>
                                            <td>
                                                <input type="button" value="Ignore" onclick="$('#NewQBItemsSel').empty(); $('#AddDiv').hide();" />
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
                <td style="width: 50%; vertical-align: top; border-left-width: thin; border-left-style: solid; padding: 20px;">
                    <table style="width: 100%">
                        <tr>
                            <td>
                                <span style="font-size: xx-large">Upload Square File</span>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <ajaxToolkit:AjaxFileUpload ID="AjaxFileUpload2" runat="server" OnUploadComplete="AjaxFileUpload2_UploadComplete"
                                    OnClientUploadComplete="SQI" MaximumNumberOfFiles="1" AllowedFileTypes="csv" />
                            </td>
                        </tr>
                        <tr>
                            <td id="SQCell"></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <table style="width: 50%; margin: auto">
                        <tr>
                            <td>
                                <span style="font-size: xx-large">Bank File</span>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <ajaxToolkit:AjaxFileUpload ID="AjaxFileUpload3" runat="server" OnUploadComplete="AjaxFileUpload3_UploadComplete"
                                    OnClientUploadComplete="bf_proc" MaximumNumberOfFiles="1" AllowedFileTypes="csv" />
                            </td>
                        </tr>
                        <tr>
                            <td id="bank_cell"></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <ajaxToolkit:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server" EnablePageMethods="True">
        </ajaxToolkit:ToolkitScriptManager>
    </form>
</body>
</html>
