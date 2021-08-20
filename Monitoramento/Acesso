function permissaoProfessores() {
  let aux = ss.getSheetByName("Escola").getRange("G8:S32").getValues();
  let valores = []
  for(let i =0; i<25;i++){
    if (aux[i][0] != ""){
      let auxi = [aux[i][0]];
      if (aux[i][8]  != ""){auxi.push(aux[i][8])}
      if (aux[i][9]  != ""){auxi.push(aux[i][9])}
      if (aux[i][10] != ""){auxi.push(aux[i][10])}
      if (aux[i][11] != ""){auxi.push(aux[i][11])}
      if (aux[i][12] != ""){auxi.push(aux[i][12])}
      valores.push(auxi);
    }
  }
  return(valores)
}
const tail = ([, ...t]) => t;
function protegerTurmas(range){
  let permProfs = permissaoProfessores();
  let professores = getProfessores();
  for(let i=0;i<permProfs.length;i++){
    let unprotected = []
    for(let j=0;j<range.length;j++){
      unprotected.push(ss.getSheetByName(permProfs[i][0]).getRange(range[j]))
    }
    let aba = ss.getSheetByName(permProfs[i][0]);
    let aux = aba.protect().setDescription('Proteção ' + permProfs[i][0]).setUnprotectedRanges(unprotected);
    aux.removeEditors(professores);
    for(let j=0;j<range.length;j++){
      let aux2 = aba.getRange(range[j]).protect().setDescription('Proteção ' + permProfs[i][0] + ' ' + range[j]);
      aux2.addEditors(tail(permProfs[i]));
    }
  }
  ss.toast("Protegido")
}
function teste(){
  let professores = getProfessores();
//  protegerTurmas(getIntevalos());
  protegerDemaisAbas();
  function protegerDemaisAbas(){
    let abas = ['GERAL','Gráficos','Aux','Escola'] 
    for(let i=0;i<abas.length;i++){
      let aux = ss.getSheetByName(abas[i]).protect().setDescription('Proteção ' + abas[i]);
      aux.removeEditors(professores);
    }
  }
  console.log("Proteger OK")
}
/*

  if (range != []){
    let unprotected = [];
    let aux = aba.protect().setDescription("Células Protegidas");
    range.forEach((_, i)=>{
      unprotected.push(aba.getRange(range[i]))
    })
/*    for (let i = 0; i < range.length; i++ ){
      unprotected.push(aba.getRange(range[i]));
    }
    aux.setUnprotectedRanges(unprotected);
    limitarAba(aba, range);
    console.log(unprotected)
    console.log(range)
  }
  aux.addEditors(getEditores());
  aux.removeEditors(getProfessores());
}*/
