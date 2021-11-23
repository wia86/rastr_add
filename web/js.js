//дожидаемся полной загрузки страницы
var GlobalVibor;
var korr_div   ;
var rgm_div    ;
var rgm_1f , rgm_nf , rgm_overload  , otkl1 , otkl2, otkl3;
var rgm_tip; // 0 текущий файл, 1 папка, 2 ПЗ в ворд по вклвдке "перегрузки"

window.onload = function () {
    GlobalVibor = 2;  /* 1 -корректировать модели Global_kor_class   2-расчитать модели Global_raschot_class */
    rgm_tip = 0;
    korr_div    = document.getElementById("div_korr");
    rgm_div     = document.getElementById("div_rgm");
    rgm_1f       =  document.getElementById("id_rgm_1f");
    rgm_nf       =  document.getElementById("id_rgm_nf");
    rgm_overload =  document.getElementById("id_rgm_overload");
    otkl1 =  document.getElementById("ID_otkl1");
    otkl2 =  document.getElementById("ID_otkl2");
    otkl3 =  document.getElementById("ID_otkl3");


    //вешаем на него событие
    rgm_1f.onclick = function() {
        document.getElementById("ID_IzFolder").style.display = "none";
        document.getElementById("id_vibor_file").style.display = "none";
        document.getElementById("ID_max_tok_save").disabled = false;
        document.getElementById("ID_gost58670").disabled = false;
        document.getElementById("ID_filtr_n2").disabled = false;
        document.getElementById("ID_otkl_ssch").disabled = false;
        document.getElementById("ID_protokol_XL").disabled = true;
        document.getElementById("ID_pz_word").disabled = false;
        document.getElementById("ID_pz_word_DIV").style.display = "none";
        document.getElementById("ID_risunok_WORD_tip").style.display = "block";
        rgm_tip = 0;
    };
    rgm_nf.onclick = function() {
        document.getElementById("ID_IzFolder").style.display = "block";
        document.getElementById("id_vibor_file").style.display = "block";
        document.getElementById("ID_max_tok_save").disabled = false;
        document.getElementById("ID_gost58670").disabled = false;
        document.getElementById("ID_filtr_n2").disabled = false;
        document.getElementById("ID_otkl_ssch").disabled = false;
        document.getElementById("ID_protokol_XL").disabled = false;
        document.getElementById("ID_pz_word").disabled = false;
        document.getElementById("ID_pz_word_DIV").style.display = "none";
        document.getElementById("ID_risunok_WORD_tip").style.display = "block";
        rgm_tip = 1;
    };
    rgm_overload.onclick = function() {    // по таблице перегрузки
        document.getElementById("ID_IzFolder").style.display = "block";
        document.getElementById("id_vibor_file").style.display = "none";
        document.getElementById("ID_ff").checked = false;
        document.getElementById("ID_max_tok_save").checked = false;
        document.getElementById("ID_max_tok_save").disabled = true;
        document.getElementById("ID_gost58670").checked = false;
        document.getElementById("ID_gost58670").disabled = true;
        document.getElementById("ID_filtr_n2").checked = false;
        document.getElementById("ID_filtr_n2").disabled = true;
        document.getElementById("ID_otkl_ssch").checked = false;
        document.getElementById("ID_otkl_ssch").disabled = true;
        document.getElementById("ID_protokol_XL").checked = false;
        document.getElementById("ID_protokol_XL").disabled = true;
        document.getElementById("ID_pz_word").disabled = true;
        document.getElementById("ID_pz_word").checked = true;
        document.getElementById("ID_pz_word_DIV").style.display = "block";
        document.getElementById("ID_risunok_WORD_tip").style.display = "none";
        rgm_tip = 2;
    };

    otkl1.onclick = function() {
        if (otkl1.checked == false){
            document.getElementById("ID_filtr_n2").checked = false;
            document.getElementById("ID_filtr_n2").disabled = true;
        }
        else {
           // getElementById("ID_filtr_n2").checked = true;
           document.getElementById("ID_filtr_n2").disabled = false;
        };

    };
};

function open_gl_div() {
    if (GlobalVibor == 1){
        GlobalVibor = 2;
        korr_div.style.display = "none";
        rgm_div.style.display = "block";
    }
    else {
        GlobalVibor = 1;
        korr_div.style.display = "block";
        rgm_div.style.display = "none";
    }
}


function open_off_ob(ell) {
    var div2 = document.getElementById(ell);
    if (div2.style.display !== "none") {
        div2.style.display = "none";}
    else {
        div2.style.display = "block"; }   /*inline */
};

function off_ob(ell) {
    var div4 = document.getElementById(ell);
    div4.style.display = "none";
 };

 function open_ob(ell) {
    var div5 = document.getElementById(ell);
    div5.style.display = "block";
 };

function open_ob_inline (ell) {
    var div2 = document.getElementById(ell);
    if (div2.style.display !== "none") {
        div2.style.display = "none";}
    else {
        div2.style.display = "inline"; }   /*inline */
};

eel.expose(js_gost);
function js_gost() {
  return 1000;
}
