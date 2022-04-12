function replicar(){
  const idModelo = "1LL6zSuhqioEqI_4lZaBfpFIjeYx7Ddf_QZ_UFBwPnNE";
  const ano = "22";

  let turmas = listaTurmas();
  console.log(listaTurmas());
  let turmasURL = [];
  for(let i=1;i<turmas.length;i++){
    console.log(turmas[i][3]); // codTurma
    let turmaID = replica(turmas[i][3], idModelo, ano);
    turmasURL.push([turmas[i][0],turmaID]);
    console.log(turmaID)
      // Alterar
      // BDt
    let turma = SpreadsheetApp.openById(turmaID);
    turma.getSheetByName("DBt").getRange("E5").setValue("20"+ano); // Ano
    turma.getSheetByName("DBt").getRange("E2").setValue([turmas[i][5]]); // Nome Professor
    turma.getSheetByName("DBt").getRange("E3").setValue([turmas[i][4]]); // Reg Professor
    turma.getSheetByName("DBt").getRange("E9").setValue([turmas[i][3]]); // COD TURMA
    turma.getSheetByName("DBt").getRange("E7").setValue([turmas[i][0]]); // Lista Piloto
      // ID
    try{
      turma.getSheetByName("ID").getRange("B6").setValue([turmas[i][5]]); // Nome Professor
      turma.getSheetByName("ID").getRange("B7").setValue([turmas[i][4]]); // Reg Professor
      turma.getSheetByName("ID").getRange("B11").setValue([turmas[i][3]]); // COD TURMA
      turma.getSheetByName("ID").getRange("B8").setValue([turmas[i][6]]); // EMAIL Professor
      turma.getSheetByName("ID").getRange("B10").setValue([turmas[i][1]]); // SED
      turma.getSheetByName("ID").getRange("B4").setValue([turmas[i][0]]); // LISTA PILOTO
      turma.getSheetByName("ID").getRange("B9").setValue([turmas[i][2]]); // TURMA
    }catch{}
      // GERAL
    try{
      turma.getSheetByName("GERAL").getRange("B2").setValue([turmas[i][2]]); // TURMA
      turma.getSheetByName("GERAL").getRange("M5").setValue(turmas[i][3].slice(2,4));
    }catch{}

    // Compartilhar
    try{turma.addEditors([turmas[i][6]]);}catch{} // Email Professor 1
    try{turma.addEditors([turmas[i][9]]);}catch{} // Email Professor 2
    try{turma.addEditors([turmas[i][11]]);}catch{} // Email Direção
    try{turma.addEditors([turmas[i][12]]);}catch{} // Email Secretaria
  }
  var folder = DriveApp.getFoldersByName("MONITORAMENTO " + anoArquivo)
  folder = folder.next();
  SpreadsheetApp.create("Turmas "+ano).getActiveSheet().getRange("A1").setValues(turmasURL);
}

function replica(codTurma, idModelo, ano) {
  // Tentando Acessar a pasta de Monitoramento do ano letivo, em caso de erro cria a pasta.
  // folder = Pasta do Ano
  try {
    var folder = DriveApp.getFoldersByName("MONITORAMENTO " + ano);
    folder = folder.next();
  }catch {
    var folder = DriveApp.createFolder("MONITORAMENTO " + ano);
  }finally{
    // Tentando Acessar a pasta da escola do ano letivo, em caso de erro cria a pasta.
    // fSchool = Pasta da escola
    try {
      var fSchool = folder.getFoldersByName(codTurma.slice(2,4));
      fSchool = fSchool.next();
    }catch {
      var fSchool = folder.createFolder(codTurma.slice(2,4));
    }finally{
      var arquivo;
      try{
        arquivo = fSchool.getFilesByName(codTurma).getId();
      }catch{
        let modelo = DriveApp.getFileById(idModelo);
        arquivo = modelo.makeCopy(codTurma, fSchool);
      }finally{
        console.log(arquivo.getId());
        console.log(msg[1] + " " + codTurma);      
        return(arquivo.getId());
      }
    }
/*
      let arq = fSchool.getFiles();
      while (arq.hasNext()){
        var file = arq.next();
        if(file != ""){
          console.log(file)
          if(file.getName()==codTurma){
            arquivo = file;
          };
        }else{
          let modelo = DriveApp.getFileById(idModelo);
          arquivo = modelo.makeCopy(codTurma, fSchool);
        }
        console.log(msg[1] + " " + codTurma);      
        return(arquivo.getUrl());
      }
    }*/
  }
}
function listaTurmas(){
//lista de turmas retirada da lista piloto "1wAdZGhlcmBoSZuccw4Ap0oOX88CUue-bYqtlEZLzckE"
  let t = SpreadsheetApp.openById("1wAdZGhlcmBoSZuccw4Ap0oOX88CUue-bYqtlEZLzckE").getSheetByName("T_Link").getRange("B2:N500").getValues();
  turmas = []
  for (let i = 0; i <t.length;i++){
    if(t[i][3]!=""){
      turmas.push(t[i])
    }
  }
/*  for(let i = 0; i<t.length;i++){
    if(t[i] != ""){
      let aux = [];
      for(let j = 0; j<t[i].length;j++){
        if(t[i][j]!=""){
          aux.push(t[i][j]);
        }
      }
      turmas.push(aux);
    }
  } */
  return turmas;
}
const msg = {
  1: "Done",
  2: "None",
}
