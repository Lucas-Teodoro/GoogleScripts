function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Menu de Funções')
    .addItem('Proteger', 'protegerAtualizado')
    .addItem('Remover', 'removerProtecoes')
    .addToUi();
}

function foo(){}

// Planilha Ativa
const ss = SpreadsheetApp.getActiveSpreadsheet();
// Lista Piloto
// const lPiloto = SpreadsheetApp.openById("1cHBMUzKly0-Stj8fl6RL3wWTwkBJUD3I4JCuM_80meA")
const lPilotoID = "1cHBMUzKly0-Stj8fl6RL3wWTwkBJUD3I4JCuM_80meA"
//const lPiloto = SpreadsheetApp.openById("10KSWydV6RtLCbTq_zqqoaMbYTuvUEhDqNr93A1IHhuo")
//Mês descrito no nome do Arquivo

function mesArquivo() {
  let mes = ss.getName().split(" ");
  return (meses[mes[2]]);
}
function anoArquivo() {
  let ano = ss.getName().split(" ");
  return ('20'+ano[0]);
}
//Dias dentro do mês
function loucuras(){
  let inicio = new Date(anoArquivo(), mesNumerico(mesArquivo())-1, 1)
  let fim    = new Date(anoArquivo(), mesNumerico(mesArquivo()), 1)
  console.log(fim)
  ss.getSheetByName("Aux").getRange("B2").setValue(mesArquivo());
  ss.getSheetByName("Aux").getRange("B1").setValue(anoArquivo());
  ss.getSheetByName("Aux").getRange("C1").setValue(fim);
}
const month = {
  1: ["JANEIRO","JAN"], 2:["FEVEREIRO","FEV"],
  3: ["MARÇO","MAR"], 4:["ABRIL","ABR"],
  5: ["MAIO","MAI"], 6:["JUNHO","JUN"],
  7: ["JULHO","JUL"], 8:["AGOSTO","AGO"],
  9: ["SETEMBRO","SET"], 10:["OUTUBRO","OUT"],
  11: ["NOVEMBRO","NOV"], 12:["DEZEMBRO","DEZ"]
}
const meses = {
  "JAN": "JANEIRO","FEV": "FEVEREIRO",
  "MAR": "MARÇO","ABR": "ABRIL",
  "MAI": "MAIO","JUN": "JUNHO",
  "JUL": "JULHO","AGO": "AGOSTO",
  "SET": "SETEMBRO","OUT": "OUTUBRO",
  "NOV": "NOVEMBRO","DEZ": "DEZEMBRO",
}
// Inserir o Menu na planilha

function proteger(){
  let sheets = ss.getSheets();
  let turmas = getTurmas();
  const filterArray = (arr1, arr2) => {
   const filtered = arr1.filter(el => {
      return arr2.indexOf(el.getName()) === -1;
   });
   return filtered;
};
  let dados = filterArray(sheets, turmas)
  for (let i = 0; i < dados.length; i++){
    protegerAba(dados[i]);
  }
  for (let i = 0; i < turmas.length; i++){
    protegerAba(ss.getSheetByName(turmas[i]), getIntervalos());
    ss.toast("Proteção turma" + turmas[i]);
  }
  ss.toast("Proteção Realizada com Sucesso")
}
function removerProtecoes(){
  let sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++){
    desprotegerAba(ss.getSheetByName(sheets[i].getName()));
  }
  ss.toast("Proteções Desativadas");
}
function getTurmas(){
  let turmas = [];
  for(let i = 1; i < 50; i++){
    if (ss.getSheetByName("Escola").getRange("F" + i).getValue() == "ENSINO FUNDAMENTAL DE 9 ANOS" || ss.getSheetByName("Escola").getRange("F" + i).getValue() == "EDUCACAO INFANTIL" || ss.getSheetByName("Escola").getRange("F" + i).getValue() == "CEE"){
      turmas.push(ss.getSheetByName("Escola").getRange("G" + i).getValue());
    }
  }
  return turmas;
}
function criarTurmas(){
  let turmas = getTurmas();
  for (let i=0; i < turmas.length;i++){
    criarTurma(turmas[i]);
  }
  Browser.msgBox("Execução de Criação de Turmas","Turmas criadas com sucesso",Browser.Buttons.OK)
}
function criarTurma(nturma){
  let base = ss.getSheetByName("Base");
  let old = ss.getSheetByName(nturma);
  /* Excluir a turma anterior antes de criar a nova */
  if (old){
    let resposta = Browser.msgBox('Excluir tabela ' + nturma,'Deseja realmente excluir a tabela ' + nturma + '? Todos os dados serão perdidos.',Browser.Buttons.YES_NO);
    console.log(resposta)
    if (resposta == "yes") {
      ss.deleteSheet(old);
      let turma = base.copyTo(ss).setName(nturma).showSheet();
      turma.getRange("I4").setValue(nturma);
    }
  }else{
  let turma = base.copyTo(ss).setName(nturma).showSheet();
  turma.getRange("I4").setValue(nturma);
  }
}
function excluirTurmas(){
  let turmas = getTurmas();
  let resposta = Browser.msgBox('Excluir turmas? ','Deseja realmente excluir todas as turmas? Todos os dados serão perdidos.',Browser.Buttons.YES_NO);
  if (resposta == "yes") {
    for (let i=0; i < turmas.length;i++){
      let old = ss.getSheetByName(turmas[i])
      if (old) ss.deleteSheet(old);
    }
  Browser.msgBox("Exclusão de Turmas","Turmas excluídas com sucesso",Browser.Buttons.OK)  
  }
}
function getLinksTurmas(){
  let turmas = getTurmas();
  let links = [];
  for (let i =0; i < turmas.length; i++){
    if(ss.getSheetByName(turmas[i])){
    let aux = ss.getUrl() + "#gid=" + ss.getSheetByName(turmas[i]).getSheetId().toString()
    links.push(aux);
    }
  }
  console.log(links)
  return links;
}
function duplicarPlaniha(){
  let nome = Browser.inputBox("Digite a Escola", "Digite o nome da Escola",Browser.Buttons.OK);
  let mes = Browser.inputBox("Digite o Mês", "Digite o mês",Browser.Buttons.OK);
  if(mesNumerico(mes)==0){Browser.msgBox("Erro","Mês digitado é inválido",Browser.Buttons.OK)}
  else{
//  ss.copy(mesNumerico(mes) + " " + mes.toToUpperCase() + " - " + nome);
  Browser.msgBox("Criar planilha da escola","Planilha da escola " + nome + " foi criada com Sucesso",Browser.Buttons.OK);
  }
}
function mesNumerico(mes){
  switch (mes.toUpperCase()){
    case "JANEIRO": return "01";
    case "FEVEREIRO": return "02";
    case "MARÇO": return "03";
    case "ABRIL": return "04";
    case "MAIO": return "05";
    case "JUNHO": return "06";
    case "JULHO": return "07";
    case "AGOSTO": return "08";
    case "SETEMBRO": return "09";
    case "OUTUBRO": return "10";
    case "NOVEMBRO": return "11";
    case "DEZEMBRO": return "12";
    default: return 0;
  }
}
function consolidar(){
  let range = ['A9:F58','G9:AP58']
  let turmas = getTurmas();
//Loop
  for(let i = 0; i < getTurmas().length;i++){
    let aba = ss.getSheetByName(turmas[i]);
  // Apagar bloqueios
    desprotegerAba(aba);
  // Sobrepor dados da aba
    sobreporDados(aba, range);
  // Proteger somente editores
  //  protegerAba(aba);
    protegerAba(aba, range);
  }
  Browser.msgBox("Consolidar planilha","Planilha consolidada com Sucesso",Browser.Buttons.OK);  
}
function foo(){
// Proteger as abas
//  for(let i = 0; i < getTurmas.length; i++){
//    protegerAba(getTurmas[i].getName, ['A9:F58','G9:AP58']);

 // }
  let aba = ss.getSheetByName("P1A")
  sobreporDados(aba, ['A9:F58','G9:AP58']);
}
function sobreporDados(aba, range = []){
  range,forEach((_,i)=>{
//  for(let i = 0; i < range.length; i++){
    console.log(aba.getName() + " ok")
    let dados = aba.getRange(range[i]).getValues();
    console.log(dados);
    aba.getRange(range[i]).setValues(dados);
    dados = []
  })
}
function protegerAba(aba, range = []){
  let aux = aba.protect().setDescription("Células Protegidas");
  if (range != []){
    let unprotected = [];
    let aux = aba.protect().setDescription("Células Protegidas");
    range.forEach((_, i)=>{
      unprotected.push(aba.getRange(range[i]))
    })
/*    for (let i = 0; i < range.length; i++ ){
      unprotected.push(aba.getRange(range[i]));
    }*/
    aux.setUnprotectedRanges(unprotected);
    limitarAba(aba, range);
    console.log(unprotected)
    console.log(range)
  }
  aux.addEditors(getEditores());
  aux.removeEditors(getProfessores());
}
function limitarAba(aba, range){
//  let aux = aba.protect().setDescription("Células Protegidas"); 
  for(let i = 0; i < range.length; i++){
    let aux = aba.getRange(range[i]).protect().setDescription("Células Protegidas");
    aux.addEditors(getEditores());
    aux.addEditors(profPorTurma(aba.getName()));
  } 
}
function profPorTurma(turma){
  for (let i = 0; i < getTurmas().length; i++){
    if (turma == ss.getSheetByName("Escola").getRange("G" + (i+8)).getValue()){
      var aux = ss.getSheetByName("Escola").getRange("O" + (i+8) + ":AA" + (i+8)).getValues() 
      return (aux[0].filter( (elem) => elem != ''))
    }
  }
}
function desprotegerAba(aba){
  let protectionsRange = aba.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  let protectionsSheet = aba.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protectionsRange.length; i++) {
    let protection = protectionsRange[i];
    if (protection.canEdit()) {
      protection.remove();
    }
  }
  for (var i = 0; i < protectionsSheet.length; i++) {
    let protection = protectionsSheet[i];
    if (protection.canEdit()) {
      protection.remove();
    }
  }
}
function getProfessores(){
  let prof = []; // lista total de professores
  let turma = ss.getSheetByName("Escola");
  let turmas = getTurmas().length;
  for (let i = 0; i < turmas; i++){
    for (let j = 0;j < 5; j++){
      let aux = turma.getRange((i+8),(j+15)).getValue();
      if (aux != ""){
        prof.push(aux);
      }
    }
  }
  return (prof.filter(function(item, pos) {
    return prof.indexOf(item) == pos;}).sort());
}
function getEditores(){
  let edit = []; // lista total de professores
  let sheet = ss.getSheetByName("Aux");
  let turmas = getTurmas().length;
  for (let i = 0; i < turmas; i++){
    let aux = sheet.getRange((i+3),5).getValue();
    if (aux != ""){
      edit.push(aux);
    }
  }
  return (edit.filter(function(item, pos) {
    return edit.indexOf(item) == pos;}).sort());
}
function getIntervalos(){
  let int = []; // lista total de intervalos a serem protegidos nas turmas
  let sheet = ss.getSheetByName("Aux");
  for (let i = 0; i < 15; i++){
    let aux = sheet.getRange((i+3),7).getValue();
    if (aux != ""){
      int.push(aux);
    }
  }
  return (int.filter(function(item, pos) {
    return int.indexOf(item) == pos;}).sort());
}
function getASheet(){
  return ss.getName();
}

function importarLPiloto(turma){
//  let turma = getASheet();
  let alunos = SpreadsheetApp.openById(lPilotoID).getSheetByName(turma).getRange("C2:K46").getValues();
  return alunos;
}
//const importarLPiloto = lPiloto.getSheetByName("1A").getRange("C2:K46").getValues();
function novo2(turma){
  let lista = [];
  let alunos = importarLPiloto(turma);
//  console.log(alunos)
  for(let i = 0; i < alunos.length;i++){
    if(alunos[i][1] != ''){
//      lista.push([alunos[i][0],alunos[i][1],alunos[i][2],alunos[i][5],alunos[i][6]])
      lista.push([alunos[i][0],alunos[i][1],alunos[i][2],alunos[i][5],alunos[i][6],alunos[i][8]])
    }    
  }
//  console.log(lista)
  return lista;
}

function dadosAlunosNaCelula(turma){
//  let turma = "1C"
  let dadosAlunos = novo2(turma);
//  console.log(dadosAlunos.length)
  ss.getSheetByName(turma).getRange("A9:F"+(9+dadosAlunos.length-1)).setValues(dadosAlunos);
//  ss.getSheetByName(turma).getRange("A9:E"+(9+dadosAlunos.length-1)).setValues(dadosAlunos);
//  console.log("Done")
}
function atualizarDados(){
  let turmas = getTurmas();
  for(let i=0;i<turmas.length;i++){
    dadosAlunosNaCelula(turmas[i]);
    console.log('Done - '+ turmas[i]);
    ss.toast('Done - '+ turmas[i]);
  }
    ss.toast('Done')
}
