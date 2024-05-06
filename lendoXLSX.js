const XLSX = require('xlsx');

const processarTabela=(nomeTabela,dadosTabela,nomeColunas)=>{
    const objetos  = dadosTabela.map(linha=>{
       const objeto={}
        for ( let i = 0; i<linha.length; i++ ){
            objeto [ nomeColunas[i] ] = linha[i]
        }
        return objeto
    })
   return objetos
}

let dados //Recebe o valor da função processar tabela
const lerPlanilha =()=> {
    const workbook = XLSX.readFile("relatorio_conf_discente_2022_novo.xlsx");
    //Buscar o nome de todas as tabelas
    const sheetNames = workbook.SheetNames;
    // sheetNames=['Docentes']


    for (const sheetName of sheetNames){
        const worksheet = workbook.Sheets[sheetName];
        const dadosTabela = XLSX.utils.sheet_to_json(worksheet,{ header: 1, blankrows: false,})
        const colunas = dadosTabela[0].map((r) => (r ));//Nome das colunas
        
        // console.log(dadosTabela)

        dados = processarTabela(sheetName,dadosTabela.slice(1),colunas)

        //Docentes
        if(sheetName==='Docentes'){
            const Instructors = dados.map(obj => {
                const novoObj = {};
                novoObj.lattes_resume_number= obj["Número do currículo Lattes"];
                novoObj.name = obj["Nome Docente"];
                novoObj.qualification = obj.Titulação;
               
                return novoObj;
              });
              
              console.log(Instructors);
        }

        //Discentes
        if(sheetName==='Discentes'){
            const students = dados.map(obj => {
                const novoObj = {};
                novoObj.name= obj["Nome Discente"];
                novoObj.student_level = obj["Nome Docente"];
                novoObj.status = obj["Situação Discente"];
                novoObj.enrollment_date = obj["Data Matrícula"];
               
                return novoObj;
              });
              
              console.log(students);
        }
        // TCC
        if(sheetName==='TCC'){
            const oritentations = dados.map(obj => {
                const novoObj = {};
                novoObj.instructor_id= obj["Identificador da Pessoa do Orientador"];
                novoObj.instructor_name = obj["Nome do Orientador"];
                if( obj["Principal?"] ==="Sim"){
                    novoObj.orientation_type = "ORIENTADOR_PRINCIPAL";
                }else{
                    novoObj.orientation_type = "CO_ORIENTADOR";

                }
                
                novoObj.completed = obj["Data fim da orientação"];//Se tiver data de fim está concluído caso não tenha não está concluído
                novoObj.nature = "Trabalho de Conclusão de Curso";
                novoObj.title = obj["Nome do Trabalho de Conclusão"];
                novoObj.year = obj["Data início da orientação"];
                novoObj.student_id = obj["Identificador da Pessoa do Autor"];

               
                return novoObj;
              });
              
              console.log(oritentations);
        }
        // Prouções Tem tanto de doscentes quanto de discentes
        if(sheetName==='Produções - Autores'){
            const productions = dados.map(obj => {
                const novoObj = {};
                novoObj.instructor_id = obj["ID Pessoa do Autor"];
                novoObj.production_type = obj["Tipo de Produção"];
                novoObj.title = obj["Nome da Produção"];
                novoObj.categorizing_author = obj["Categoria do Autor"];
                novoObj.year = obj["Ano da Produção"];
                novoObj.purpose = obj["Projeto"];
                novoObj.research_line = obj["Linha de Pesquisa"];


               
                return novoObj;
              });
              
              console.log(productions);
        }


        
    
    }

}


lerPlanilha();
