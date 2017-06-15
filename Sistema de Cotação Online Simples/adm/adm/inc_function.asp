<%
'Função para tamanho de arquivo
Function fun_tamanhos(auxtamanho)
	if auxtamanho < 1024 then
		fun_tamanhos = auxtamanho & " BY"
	else
		if auxtamanho > 1024 and auxtamanho < 1048576 then
			fun_tamanhos = formatnumber((auxtamanho/1024),3) & " KB"
		else
			if auxtamanho > 1048576 and auxtamanho < 1099511627776 then
				fun_tamanhos = formatnumber((auxtamanho/1048576),3) & " MB"
			else
				fun_tamanhos = formatnumber((auxtamanho/1099511627776),3) & " GB"
			end if
		end if
	end if
End Function

Function fun_tamanhos2(auxtamanho,auxtamanho2)
	if auxtamanho2 = "BY" then
		fun_tamanhos2 = auxtamanho		
	else
		if auxtamanho2 = "KB" then
			fun_tamanhos2 = auxtamanho * 1024
		else
			if auxtamanho2 = "MB" then
				fun_tamanhos2 = auxtamanho * 1048576
			else
				fun_tamanhos2 = auxtamanho * 1099511627776
			end if
		end if
	end if
End Function	

'---------------------------------
'DIRETÓRIO PADRÃO DE ARMAZENAMENTO
arquivo = ""
'---------------------------------

if topo <> "sim" then
%>

<script language="javascript" type="text/javascript">
function setPointer(theRow, thePointerColor)
{
    if (thePointerColor == '' || typeof(theRow.style) == 'undefined') {
        return false;
    }
    if (typeof(document.getElementsByTagName) != 'undefined') {
        var theCells = theRow.getElementsByTagName('td');
    }
    else if (typeof(theRow.cells) != 'undefined') {
        var theCells = theRow.cells;
    }
    else {
        return false;
    }

    var rowCellsCnt  = theCells.length;
    for (var c = 0; c < rowCellsCnt; c++) {
        theCells[c].style.backgroundColor = thePointerColor;
    }

    return true;
} // end of the 'setPointer()' function

/***
* Descrição.: formata um campo do formulário de
* acordo com a máscara informada...
* Parâmetros: - objForm (o Objeto Form)
* - strField (string contendo o nome
* do textbox)
* - sMask (mascara que define o
* formato que o dado será apresentado,
* usando o algarismo "9" para
* definir números e o símbolo "!" para
* qualquer caracter...
* - evtKeyPress (evento)
* Uso.......: <input type="textbox"
* name="xxx".....
* onkeypress="return txtBoxFormat(document.rcfDownload, 'str_cep', '99999-999', event);">
* Observação: As máscaras podem ser representadas como os exemplos abaixo:
* CEP -> 99.999-999
* CPF -> 999.999.999-99
* CNPJ -> 99.999.999/9999-99
* Data -> 99/99/9999
* Tel Resid -> (99) 999-9999
* Tel Cel -> (99) 9999-9999
* Processo -> 99.999999999/999-99
* C/C -> 999999-!
* E por aí vai...
***/

function txtBoxFormat(strField, sMask, evtKeyPress) {
    var i, nCount, sValue, fldLen, mskLen,bolMask, sCod, nTecla;

    if(window.event) { // Internet Explorer
      nTecla = evtKeyPress.keyCode; }
    else if(evtKeyPress.which) { // Nestcape / firefox
      nTecla = evtKeyPress.which;
    }
//se for backspace não faz nada
if (nTecla != 8){
    sValue = document.getElementById(strField).value;
// alert(sValue);

    // Limpa todos os caracteres de formatação que
    // já estiverem no campo.
    sValue = sValue.toString().replace( "-", "" );
    sValue = sValue.toString().replace( "-", "" );
    sValue = sValue.toString().replace( ".", "" );
    sValue = sValue.toString().replace( ".", "" );
    sValue = sValue.toString().replace( ":", "" );
    sValue = sValue.toString().replace( ":", "" );
    sValue = sValue.toString().replace( "/", "" );
    sValue = sValue.toString().replace( "/", "" );
    sValue = sValue.toString().replace( "(", "" );
    sValue = sValue.toString().replace( "(", "" );
    sValue = sValue.toString().replace( ")", "" );
    sValue = sValue.toString().replace( ")", "" );
    sValue = sValue.toString().replace( " ", "" );
    sValue = sValue.toString().replace( " ", "" );
    fldLen = sValue.length;
    mskLen = sMask.length;

    i = 0;
    nCount = 0;
    sCod = "";
    mskLen = fldLen;

    while (i <= mskLen) {
      bolMask = ((sMask.charAt(i) == "-") || (sMask.charAt(i) == ".")  || (sMask.charAt(i) == ":") || (sMask.charAt(i) 

== "/"))
      bolMask = bolMask || ((sMask.charAt(i) == "(") || (sMask.charAt(i) == ")") || (sMask.charAt(i) == " "))

      if (bolMask) {
        sCod += sMask.charAt(i);
        mskLen++; }
      else {
        sCod += sValue.charAt(nCount);
        nCount++;
      }

      i++;
    }

    document.getElementById(strField).value = sCod;

    if (nTecla != 8) { // backspace
      if (sMask.charAt(i-1) == "9") { // apenas números...
        return ((nTecla > 47) && (nTecla < 58)); } // números de 0 a 9
      else { // qualquer caracter...
        return true;
      } }
    else {
      return true;
    }
}//fim do if que verifica se é backspace
}
//Fim da Função Máscaras Gerais


fCol='444444'; //face colour.
sCol='FF0000'; //seconds colour.
mCol='444444'; //minutes colour.
hCol='444444'; //hours colour.

YCbase=30; //Clock height.
XCbase=30; //Clock width.

H='...';
H=H.split('');
M='....';
M=M.split('');
S='.....';
S=S.split('');
NS4=(document.layers);
NS6=(document.getElementById&&!document.all);
IE4=(document.all);
Ypos=0;
Xpos=0;
cdots=12;
Split=360/cdots;
if (NS6){
for (i=1; i < cdots+1; i++){
document.write('<div id="n6Digits'+i+'" style="position:absolute;top:0px;left:0px;width:30px;height:30px;font-family:Arial;font-size:10px;color:#'+fCol+';text-align:center;padding-top:10px">'+i+'</div>');
}
for (i=0; i < M.length; i++){
document.write('<div id="Ny'+i+'" style="position:absolute;top:0px;left:0px;width:2px;height:2px;font-size:2px;background:#'+mCol+'"></div>');
}
for (i=0; i < H.length; i++){
document.write('<div id="Nz'+i+'" style="position:absolute;top:0px;left:0px;width:2px;height:2px;font-size:2px;background:#'+hCol+'"></div>');
}
for (i=0; i < S.length; i++){
document.write('<div id="Nx'+i+'" style="position:absolute;top:0px;left:0px;width:2px;height:2px;font-size:2px;background:#'+sCol+'"></div>');
}
}
if (NS4){
dgts='1 2 3 4 5 6 7 8 9 10 11 12';
dgts=dgts.split(' ')
for (i=0; i < cdots; i++){
document.write('<layer name=nsDigits'+i+' top=0 left=0 height=30 width=30><center><font face=Arial size=1 color='+fCol+'>'+dgts[i]+'</font></center></layer>');
}
for (i=0; i < M.length; i++){
document.write('<layer name=ny'+i+' top=0 left=0 bgcolor='+mCol+' clip="0,0,2,2"></layer>');
}
for (i=0; i < H.length; i++){
document.write('<layer name=nz'+i+' top=0 left=0 bgcolor='+hCol+' clip="0,0,2,2"></layer>');
}
for (i=0; i < S.length; i++){
document.write('<layer name=nx'+i+' top=0 left=0 bgcolor='+sCol+' clip="0,0,2,2"></layer>');
}
}
if (IE4){
document.write('<div style="position:absolute;top:0px;left:0px"><div style="position:relative">');
for (i=1; i < cdots+1; i++){
document.write('<div id="ieDigits" style="position:absolute;top:0px;left:0px;width:30px;height:30px;font-family:Arial;font-size:10px;color:'+fCol+';text-align:center;padding-top:10px">'+i+'</div>');
}
document.write('</div></div>')
document.write('<div style="position:absolute;top:0px;left:0px"><div style="position:relative">');
for (i=0; i < M.length; i++){
document.write('<div id=y style="position:absolute;width:2px;height:2px;font-size:2px;background:'+mCol+'"></div>');
}
document.write('</div></div>')
document.write('<div style="position:absolute;top:0px;left:0px"><div style="position:relative">');
for (i=0; i < H.length; i++){
document.write('<div id=z style="position:absolute;width:2px;height:2px;font-size:2px;background:'+hCol+'"></div>');
}
document.write('</div></div>')
document.write('<div style="position:absolute;top:0px;left:0px"><div style="position:relative">');
for (i=0; i < S.length; i++){
document.write('<div id=x style="position:absolute;width:2px;height:2px;font-size:2px;background:'+sCol+'"></div>');
}
document.write('</div></div>')
}



function clock(){
time = new Date ();
secs = time.getSeconds();
sec = -1.57 + Math.PI * secs/30;
mins = time.getMinutes();
min = -1.57 + Math.PI * mins/30;
hr = time.getHours();
hrs = -1.57 + Math.PI * hr/6 + Math.PI*parseInt(time.getMinutes())/360;

if (NS6){
Ypos=window.pageYOffset+window.innerHeight-YCbase-25;
Xpos=window.pageXOffset+window.innerWidth-XCbase-30;
for (i=1; i < cdots+1; i++){
 document.getElementById("n6Digits"+i).style.top=Ypos-15+YCbase*Math.sin(-1.56 +i *Split*Math.PI/180)
 document.getElementById("n6Digits"+i).style.left=Xpos-15+XCbase*Math.cos(-1.56 +i*Split*Math.PI/180)
 }
for (i=0; i < S.length; i++){
 document.getElementById("Nx"+i).style.top=Ypos+i*YCbase/4.1*Math.sin(sec);
 document.getElementById("Nx"+i).style.left=Xpos+i*XCbase/4.1*Math.cos(sec);
 }
for (i=0; i < M.length; i++){
 document.getElementById("Ny"+i).style.top=Ypos+i*YCbase/4.1*Math.sin(min);
 document.getElementById("Ny"+i).style.left=Xpos+i*XCbase/4.1*Math.cos(min);
 }
for (i=0; i < H.length; i++){
 document.getElementById("Nz"+i).style.top=Ypos+i*YCbase/4.1*Math.sin(hrs);
 document.getElementById("Nz"+i).style.left=Xpos+i*XCbase/4.1*Math.cos(hrs);
 }
}
if (NS4){
Ypos=window.pageYOffset+window.innerHeight-YCbase-20;
Xpos=window.pageXOffset+window.innerWidth-XCbase-30;
for (i=0; i < cdots; ++i){
 document.layers["nsDigits"+i].top=Ypos-5+YCbase*Math.sin(-1.045 +i*Split*Math.PI/180)
 document.layers["nsDigits"+i].left=Xpos-15+XCbase*Math.cos(-1.045 +i*Split*Math.PI/180)
 }
for (i=0; i < S.length; i++){
 document.layers["nx"+i].top=Ypos+i*YCbase/4.1*Math.sin(sec);
 document.layers["nx"+i].left=Xpos+i*XCbase/4.1*Math.cos(sec);
 }
for (i=0; i < M.length; i++){
 document.layers["ny"+i].top=Ypos+i*YCbase/4.1*Math.sin(min);
 document.layers["ny"+i].left=Xpos+i*XCbase/4.1*Math.cos(min);
 }
for (i=0; i < H.length; i++){
 document.layers["nz"+i].top=Ypos+i*YCbase/4.1*Math.sin(hrs);
 document.layers["nz"+i].left=Xpos+i*XCbase/4.1*Math.cos(hrs);
 }
}

if (IE4){
Ypos=document.body.scrollTop+window.document.body.clientHeight-YCbase-20;
Xpos=document.body.scrollLeft+window.document.body.clientWidth-XCbase-20;
for (i=0; i < cdots; ++i){
 ieDigits[i].style.pixelTop=Ypos-15+YCbase*Math.sin(-1.045 +i *Split*Math.PI/180)
 ieDigits[i].style.pixelLeft=Xpos-15+XCbase*Math.cos(-1.045 +i *Split*Math.PI/180)
 }
for (i=0; i < S.length; i++){
 x[i].style.pixelTop =Ypos+i*YCbase/4.1*Math.sin(sec);
 x[i].style.pixelLeft=Xpos+i*XCbase/4.1*Math.cos(sec);
 }
for (i=0; i < M.length; i++){
 y[i].style.pixelTop =Ypos+i*YCbase/4.1*Math.sin(min);
 y[i].style.pixelLeft=Xpos+i*XCbase/4.1*Math.cos(min);
 }
for (i=0; i < H.length; i++){
 z[i].style.pixelTop =Ypos+i*YCbase/4.1*Math.sin(hrs);
 z[i].style.pixelLeft=Xpos+i*XCbase/4.1*Math.cos(hrs);
 }
}
setTimeout('clock()',100);
}
clock();
//-->
</script>
<%else%>
<script language="javascript" type="text/javascript">
function setPointer(theRow, thePointerColor)
{
    if (thePointerColor == '' || typeof(theRow.style) == 'undefined') {
        return false;
    }
    if (typeof(document.getElementsByTagName) != 'undefined') {
        var theCells = theRow.getElementsByTagName('td');
    }
    else if (typeof(theRow.cells) != 'undefined') {
        var theCells = theRow.cells;
    }
    else {
        return false;
    }

    var rowCellsCnt  = theCells.length;
    for (var c = 0; c < rowCellsCnt; c++) {
        theCells[c].style.backgroundColor = thePointerColor;
    }

    return true;
} // end of the 'setPointer()' function

/***
* Descrição.: formata um campo do formulário de
* acordo com a máscara informada...
* Parâmetros: - objForm (o Objeto Form)
* - strField (string contendo o nome
* do textbox)
* - sMask (mascara que define o
* formato que o dado será apresentado,
* usando o algarismo "9" para
* definir números e o símbolo "!" para
* qualquer caracter...
* - evtKeyPress (evento)
* Uso.......: <input type="textbox"
* name="xxx".....
* onkeypress="return txtBoxFormat(document.rcfDownload, 'str_cep', '99999-999', event);">
* Observação: As máscaras podem ser representadas como os exemplos abaixo:
* CEP -> 99.999-999
* CPF -> 999.999.999-99
* CNPJ -> 99.999.999/9999-99
* Data -> 99/99/9999
* Tel Resid -> (99) 999-9999
* Tel Cel -> (99) 9999-9999
* Processo -> 99.999999999/999-99
* C/C -> 999999-!
* E por aí vai...
***/

function txtBoxFormat(strField, sMask, evtKeyPress) {
    var i, nCount, sValue, fldLen, mskLen,bolMask, sCod, nTecla;

    if(window.event) { // Internet Explorer
      nTecla = evtKeyPress.keyCode; }
    else if(evtKeyPress.which) { // Nestcape / firefox
      nTecla = evtKeyPress.which;
    }
//se for backspace não faz nada
if (nTecla != 8){
    sValue = document.getElementById(strField).value;
// alert(sValue);

    // Limpa todos os caracteres de formatação que
    // já estiverem no campo.
    sValue = sValue.toString().replace( "-", "" );
    sValue = sValue.toString().replace( "-", "" );
    sValue = sValue.toString().replace( ".", "" );
    sValue = sValue.toString().replace( ".", "" );
    sValue = sValue.toString().replace( ":", "" );
    sValue = sValue.toString().replace( ":", "" );
    sValue = sValue.toString().replace( "/", "" );
    sValue = sValue.toString().replace( "/", "" );
    sValue = sValue.toString().replace( "(", "" );
    sValue = sValue.toString().replace( "(", "" );
    sValue = sValue.toString().replace( ")", "" );
    sValue = sValue.toString().replace( ")", "" );
    sValue = sValue.toString().replace( " ", "" );
    sValue = sValue.toString().replace( " ", "" );
    fldLen = sValue.length;
    mskLen = sMask.length;

    i = 0;
    nCount = 0;
    sCod = "";
    mskLen = fldLen;

    while (i <= mskLen) {
      bolMask = ((sMask.charAt(i) == "-") || (sMask.charAt(i) == ".")  || (sMask.charAt(i) == ":") || (sMask.charAt(i) 

== "/"))
      bolMask = bolMask || ((sMask.charAt(i) == "(") || (sMask.charAt(i) == ")") || (sMask.charAt(i) == " "))

      if (bolMask) {
        sCod += sMask.charAt(i);
        mskLen++; }
      else {
        sCod += sValue.charAt(nCount);
        nCount++;
      }

      i++;
    }

    document.getElementById(strField).value = sCod;

    if (nTecla != 8) { // backspace
      if (sMask.charAt(i-1) == "9") { // apenas números...
        return ((nTecla > 47) && (nTecla < 58)); } // números de 0 a 9
      else { // qualquer caracter...
        return true;
      } }
    else {
      return true;
    }
}//fim do if que verifica se é backspace
}
//Fim da Função Máscaras Gerais



//-->
</script>
<%end if%>