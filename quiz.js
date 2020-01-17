/*   Riccardo Bertossa, (c) 2017, 2018, 2019, 2020
 *   
 *   Crea e gestisce i test degli arbitri
 *   genera automaticamente il questionario e la presentazione con le domande e le risposte
 *   genera automaticamente il questionario cartaceo 
 *   implementa un sistema di valutazione della domanda
*/


function calcNline_approx(str) {
        return Math.floor(str.length / 80) +1
}

function prepare_new_quiz_from_selected(){
  var question_bank_ID = 'INSERT ID'
  var sscq = SpreadsheetApp.openById(question_bank_ID);
  var sheetcq = sscq.getSheets()[0];
  var testcq = sscq.getSheets()[sscq.getSheets().length-1];
  var color_selection='#ffe599'
  
  var id_domande = new Array();
  var corrette = new Array();
  
  //trova gli id delle domande selezionate
  var n_domande_database=0
  for (var i=2;;i++){
    if(sheetcq.getSheetValues(i,1, 1, 1)[0][0]!==''){
      n_domande_database++;
    } else {
      break;
    }
    
    //se il colore della selezione è quello corretto, aggiungi l'id alla lista degli id.
    
    if (sheetcq.getRange(i,2, 1, 1).getBackground()===color_selection){
      corrette.push(new Array());
      id_domande.push(sheetcq.getSheetValues(i,1, 1, 1)[0][0]);
              for(var jd=7;;jd++){
          if (sheetcq.getSheetValues(i,jd, 1, 1)[0][0]!=='') {
            if (sheetcq.getRange(i,jd, 1, 1).getBackground()==='#00ff00') {
              corrette[corrette.length-1].push(true);
            } else {
              corrette[corrette.length-1].push(false);
            }
          } else {
            break;
          }
        }
    }
  }
  
  
  Logger.log(id_domande);
  
  //scrive gli id nel foglio delle domande
  
  for (var i=0;i<id_domande.length;i++){
        testcq.getRange(i+5, 1).setValue(id_domande[i]);
        for (var j=0;j<corrette[i].length;j++) {
           if (corrette[i][j])
            testcq.getRange(i+5, 3+j).setValue(1);
           else
            testcq.getRange(i+5, 3+j).setValue(-10);
        }
  }

}

function get_question_from_database(sheetcq,  ///oggetto che decrive il particolare foglio. Ad esempio può essere recuperato con SpreadsheetApp.openById(question_bank_ID).getSheets()[0]
                                    n_domande_database,
                                    id_domande ///id della domanda
                                    ){
// restituisce [ testo_domanda, [array_risposte],[array_risposta_è_vera]]
for(var j=2;j<=n_domande_database+1;j++){
      if(sheetcq.getSheetValues(j,1, 1, 1)[0][0]===id_domande){
        //trova le risposte
        var risposte= new Array();
        var corretta= new Array();
        for(var jd=7;;jd++){
          if (sheetcq.getSheetValues(j,jd, 1, 1)[0][0]!=='') {
            risposte.push(sheetcq.getSheetValues(j,jd, 1, 1)[0][0]);
            if (sheetcq.getRange(j,jd, 1, 1).getBackground()==='#00ff00') {
              corretta.push(true);
            } else {
              corretta.push(false);
            }
          } else {
            break;
          }
        }
        return [sheetcq.getSheetValues(j,2, 1, 1)[0][0],risposte,corretta];
      }
    }
    return ["Questa frase è falsa",["vero","falso"],[false,false]];
}


function make_a_new_quiz() {


  
  var question_bank_ID = 'INSERT ID'
 
  
  var quiz_template_ID = 'INSERT ID'
  var quiz_scritto_template_ID = 'INSERT ID';
  var sscq = SpreadsheetApp.openById(question_bank_ID);

  var sheetcq = sscq.getSheets()[0];
  var testcq = sscq.getSheets()[sscq.getSheets().length-1];
  var title = testcq.getSheetValues(1,1,1,1)[0][0];
  var description = testcq.getSheetValues(2,1,1,1)[0][0];
  var data_test=testcq.getSheetValues(1,2,1,1)[0][0];
  // trova il numero di domande del test
  
  var id_domande = new Array();
  for (var i=5;;i++){
    if(testcq.getSheetValues(i,1, 1, 1)[0][0]!==''){
      id_domande.push(testcq.getSheetValues(i,1, 1, 1)[0][0]);
    } else {
      break;
    }
  }
  
  // trova il numero di domande del database
  var n_domande_database=0
  for (var i=2;;i++){
     if(sheetcq.getSheetValues(i,1, 1, 1)[0][0]!==''){
      n_domande_database++;
    } else {
      break;
    }
  }
  
  Logger.log(id_domande.length);
  Logger.log(n_domande_database);
  
  Logger.log(id_domande);
  
  var questions_answers = new Array();
  //popola le domande e le risposte dalla lista delle domande
  for(var i=0;i<id_domande.length;i++) {

      questions_answers.push(
            get_question_from_database(sheetcq, n_domande_database,id_domande[i])
      );
  
  }
  
  Logger.log(title);
  Logger.log(description);
  Logger.log(questions_answers);
  
  var slides=SlidesApp.create(title);
  var master_quiz = DriveApp.getFileById(quiz_template_ID);
  var master_quiz_scritto = DriveApp.getFileById(quiz_scritto_template_ID);
 
 var quiz_scritto = master_quiz_scritto.makeCopy(title);
 var test_scritto = DocumentApp.openById(quiz_scritto.getId());
 
 test_scritto.getBody().replaceText("%TITOLOTEST%", title);
 test_scritto.getBody().replaceText("%DESCRIZIONETEST%", description);
 test_scritto.getBody().replaceText("%DATATEST%",data_test );
  
   var quiz = master_quiz.makeCopy(title);
   var id = quiz.getId();
    
   var form = FormApp.openById(id);
    
    form.setTitle(title)     
     .setDescription(description)
     .setConfirmationMessage('Grazie per aver risposto!');
 
    var listid='';
    var listD;
  for (var i=0;i<questions_answers.length;i++) {
    // startRow, startColumn, numRows, numColumns
    var text = questions_answers[i][0];
    var options = questions_answers[i][1];
    var corretta = questions_answers[i][2];
    
    
    Logger.log(text);
   
    { // Checkbox
      var item = form.addCheckboxItem();
      item.setTitle(text).setPoints(1); // non sono sicuro fosse item
      
      //slide
      var current_slide2=slides.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      var current_slide=slides.appendSlide(SlidesApp.PredefinedLayout.BLANK);
            
      
      
      DX=slides.getPageWidth();
      DY=slides.getPageHeight();
      
      //with answer
      var question_text=current_slide.insertTextBox(text);
      question_text.setWidth(DX);
      question_text.setHeight(40);
      var h_question=question_text.getInherentHeight()*question_text.getTransform().getScaleY()*calcNline_approx(text);
      question_text.getText().getTextStyle().setFontSize(20);
      var question_answers=current_slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,0,h_question,DX,40);
      //without answer
      var question_text2=current_slide2.insertTextBox(text);
      question_text2.setWidth(DX)
      question_text2.setHeight(40);
      question_text2.getText().getTextStyle().setFontSize(20);
      var question_answers2=current_slide2.insertShape(SlidesApp.ShapeType.TEXT_BOX,0,h_question,DX,40);
      
      
      //var slide_elements=slides.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY).getPageElements();
      //slide_elements[0].asShape().getText().setText(text).getTextStyle().setFontSize(20);
      
            
      
     
     //////////////////
      
      // versione cartacea
      test_scritto.getBody().appendParagraph('');
      var domanda=test_scritto.getBody().appendListItem(text);
      if (listid==='') {
        listid=domanda.getListId();
        listD=domanda;
      }
      domanda.setListId(listD);
      
    var choices = new Array();
    
    var testo_risposta="";
    var listidR='';
    var listR;
    for (var j in options) {
      choices.push(item.createChoice(options[j],corretta[j]));
      if (corretta[j])
         testo_risposta+="_"+options[j]+"\n";
      else
         testo_risposta+=options[j]+"\n";

      //slides
      var a_capo='\n';
      if (j == options.length-1 )
         a_capo='';
      var app_txt=question_answers.getText().appendText(options[j]+a_capo);
      var app_txt2=question_answers2.getText().appendText(options[j]+a_capo);
      if (corretta[j])
          app_txt.getTextStyle().setBold(true).setUnderline(true);
      else 
          app_txt.getTextStyle().setBold(false).setUnderline(false);

      var rispostaR=test_scritto.getBody().appendListItem(options[j]).setNestingLevel(1).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
      if (listidR==='') {
        listidR=rispostaR.getListId();
        listR=rispostaR;
      }
      rispostaR.setListId(listR);
    }
      item.setChoices(choices);
      //slide_elements[1].asShape().getText().setText(testo_risposta).getTextStyle().setFontSize(18);
      question_answers.getText().getListStyle().applyListPreset(SlidesApp.ListPreset.DIGIT_NESTED);
      question_answers.getText().getTextStyle().setFontSize(20).setForegroundColor('#595959');
      question_answers2.getText().getListStyle().applyListPreset(SlidesApp.ListPreset.DIGIT_NESTED);
      question_answers2.getText().getTextStyle().setFontSize(20).setForegroundColor('#595959');


    }
    
    
  }
}

function slide_test() {

      var current_slides=SlidesApp.openById('INSERT ID')
      var current_slide=current_slides.getSlides()[0];
      
      
      DX=current_slides.getPageWidth();
      DY=current_slides.getPageHeight();
      var question_text=current_slide.insertTextBox('Domanda di test bla bla bla. Domanda di test bla bla bla. Domanda di test bla bla bla. Domanda di test bla bla bla. Domanda di test bla bla bla. Domanda di test bla bla bla. Domanda di test bla bla bla. Domanda di test bla bla bla. ');
      //question_text.alignOnPage(SlidesApp.AlignmentPosition.CENTER)
      question_text.setWidth(DX)
      var h_question=question_text.getInherentHeight()*question_text.getTransform().getScaleY();
      question_text.getText().getTextStyle().setFontSize(20);
      var question_answers=current_slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,0,h_question,DX,40);
      
      
      //var slide_elements=slides.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY).getPageElements();
      //slide_elements[0].asShape().getText().setText(text).getTextStyle().setFontSize(20);
      
      question_answers.getText().appendText('Risposta 1\n')
            .appendText('Risposta 2. infatti pippo faceva le puzzette\n')
            .appendText('Risposta 3. ma quante diavolo di risposte ha questa domanda di prova? questa sembra anche parecchio lunga, ma a cosa serve? cosa sta dicendo?');
            
      question_answers.getText().getListStyle().applyListPreset(SlidesApp.ListPreset.DIGIT_NESTED);
      question_answers.getText().getTextStyle().setFontSize(20).setForegroundColor('#595959');
      
      question_answers.getText().getParagraphs()[1].getRange().getTextStyle().setBold(true).setUnderline(true);
}

function write_statistics(){
/*
 * Questa funzione scrive nel foglio delle domande la percentuali di errori (utile per classificare la difficoltà delle domande)
 * ricordarsi di modificare quiz_id
*/


  
  var question_bank_ID = 'INSERT ID'
     var quiz_id= 'INSERT ID';
  
  var sscq = SpreadsheetApp.openById(question_bank_ID);

  var sheetcq = sscq.getSheets()[0];
  var testcq = sscq.getSheets()[sscq.getSheets().length-1];
  var title = testcq.getSheetValues(1,1,1,1)[0][0];
  var description = testcq.getSheetValues(2,1,1,1)[0][0];
  
  var data_test=testcq.getSheetValues(1,2,1,1)[0][0]
  // trova il numero di domande del test
  
  var id_domande = new Array();
  for (var i=5;;i++){
    if(testcq.getSheetValues(i,1, 1, 1)[0][0]!==''){
      id_domande.push(testcq.getSheetValues(i,1, 1, 1)[0][0]);
    } else {
      break;
    }
  }
  
  // trova il numero di domande del database
  var n_domande_database=0
  for (var i=2;;i++){
     if(sheetcq.getSheetValues(i,1, 1, 1)[0][0]!==''){
      n_domande_database++;
    } else {
      break;
    }
  }
  
  Logger.log(id_domande.length);
  Logger.log(n_domande_database);
  
  Logger.log(id_domande);
  
  
  
  // legge il numero di domande sbagliate e calcola la percentuale di errori
  

 
  var quiz=FormApp.openById(quiz_id);

  var tot_score_answer=[];
  var score_answer=[];
  var formResponses=quiz.getResponses();
  var questions=quiz.getItems();

  for (var j=0;j<questions.length;j++){
     tot_score_answer[j]=0;
     score_answer[j]=0;
  }
  for (var i=0;i<formResponses.length;i++){
    for (var j=0;j<questions.length;j++){
     var response = formResponses[i].getGradableResponseForItem(questions[j]);
     score_answer[j] += response.getScore();
     tot_score_answer[j]++;
    }
    
  }
  
  var questions_answers = new Array();
  //popola le domande e le risposte dalla lista delle domande
  for(var i=0;i<id_domande.length;i++) {
    for(var j=1;j<=n_domande_database;j++){
      if(sheetcq.getSheetValues(j,1, 1, 1)[0][0]===id_domande[i]){
      
      //scrive le statistiche
        sheetcq.getRange(j,3).setValue(data_test);
        sheetcq.getRange(j,4).setValue(1-score_answer[i]/tot_score_answer[i]);
        break;
      }
    }
  }
  
}

function get_mail_scores(){

 /*
  * Questa funzione compila la scheda "contatti" con i risultati per nome prendendo la colonna.
  *
  * IMPORTANTE: modificare la variabile colonna_punteggi alla colonna dove si desidera siano scritti i risultati
  * PRIMA DI ESEGUIRE LA FUNZIONE
  * modificare anche l'ID del quiz! (si può prendere dall'URL del test)
  * se necessario modificare la scheda risultati
 */

  var quiz_id='INSERT ID';
  var question_bank_ID = 'INSERT ID';
 
  var quiz=FormApp.openById(quiz_id);
  var sscq = SpreadsheetApp.openById(question_bank_ID);
  risultati=sscq.getSheetByName("Contatti");
  
  var colonna_punteggi=9;



  
  // legge le domande ed eventuali modifiche ai punteggi delle risposte (personalizzati)

  var sheetcq = sscq.getSheets()[0];
  var testcq = sscq.getSheets()[sscq.getSheets().length-1];
  var title = testcq.getSheetValues(1,1,1,1)[0][0];
  var description = testcq.getSheetValues(2,1,1,1)[0][0];
  
  var data_test=testcq.getSheetValues(1,2,1,1)[0][0]
  
  
    // trova il numero di domande del database
  var n_domande_database=0
  for (var i=2;;i++){
     if(sheetcq.getSheetValues(i,1, 1, 1)[0][0]!==''){
      n_domande_database++;
    } else {
      break;
    }
  }
  
  Logger.log(n_domande_database);
  
  
  
  // trova il numero di domande del test
  
  var idx_answer_scores = new Array(); // answer_scores[idx_answer_scores[idx_domanda]] sono i punteggi di ciascuna risposta (se idx_domanda>0)
  var answer_scores = new Array(); // punteggi di ciascuna risposta
  var answer_scores_cont=0;
  
  var id_domande = new Array();
  for (var i=5;;i++){
    if(testcq.getSheetValues(i,1, 1, 1)[0][0]!==''){
      id_domande.push(testcq.getSheetValues(i,1, 1, 1)[0][0]);
      if (testcq.getRange(i,1, 1, 1).getBackground()==='#ff9900') {
          answer_scores.push(new Array());
          idx_answer_scores.push(answer_scores_cont);
          for (var j=3;;j++){
               if(testcq.getSheetValues(i,j, 1, 1)[0][0]==='') break;
               // legge i punteggi di ciascuna risposta
               answer_scores[answer_scores_cont].push(testcq.getSheetValues(i,j, 1, 1)[0][0])
          }
          answer_scores_cont++;
      } else {
          idx_answer_scores.push(-1);
      }
    } else {
      break;
    }
  }




  //trova il numero di arbitri e osservatori
  var n_arbitri_osservatori=0
  for (var i=2;;i++){
     if(risultati.getSheetValues(i,1, 1, 1)[0][0]!==''){
      n_arbitri_osservatori++;
    } else {
      break;
    }
  }
  
  var formResponses=quiz.getResponses();
  var questions=quiz.getItems();
  for (var j=0;j<questions.length;j++){
     questions[j].asCheckboxItem().setPoints(1);
  }
  for (var i=0;i<formResponses.length;i++){
    var email=formResponses[i].getRespondentEmail();
    var totScore=0;
    for (var j=0;j<questions.length;j++){
     var response = formResponses[i].getGradableResponseForItem(questions[j]);
     if (idx_answer_scores[j]==-1){
       totScore += response.getScore();
     } else { // calcola il punteggio in base al valore dei punti di ciascuna risposta (specificata nel foglio che ha generato il test)
       answer=response.getResponse();
       // ...
       var q = get_question_from_database(sheetcq, n_domande_database,id_domande[j]);
       var score=0;
       var score_correct=0;
       for (var k=0;k<q[1].length;k++){
         if (answer_scores[idx_answer_scores[j]][k]>0) {
           score_correct+=answer_scores[idx_answer_scores[j]][k];
         }
       }
       for (var k1=0;k1<answer.length;k1++){
         for (var k=0;k<q[1].length;k++){
             if (answer[k1].trim().toLowerCase()==q[1][k].trim().toLowerCase()) {
                score+=answer_scores[idx_answer_scores[j]][k];
                break;
             }
         }
       }
       if (score != score_correct){
         score=0;
       } else {
         score=1;
       }
       response.setScore(score);
       formResponses[i].withItemGrade(response);
       totScore += score;
     }
    }
    Logger.log(email);
    Logger.log(totScore);
    // trova la riga corrispondente all'indirizzo email e aggiungi nella colonna il risultato del test
    var j=0;
    for (;j<n_arbitri_osservatori;j++){
      if (risultati.getSheetValues(j+2,3,1,1)[0][0].toLowerCase()===email.toLowerCase()){
        risultati.getRange(j+2,colonna_punteggi).setValue(totScore);
        break;
      }
    }
    if (j===n_arbitri_osservatori) {
      //cerca nella colonna mail2
        var j2=0;
        for (;j2<n_arbitri_osservatori;j2++){
          if (risultati.getSheetValues(j2+2,4,1,1)[0][0].toLowerCase()===email.toLowerCase()){
            risultati.getRange(j2+2,colonna_punteggi).setValue(totScore);
            break;
           }
        }
      if (j2===n_arbitri_osservatori){
        var nuova_riga=['?','?',email];
        for (var k=4;k<colonna_punteggi;k++)
           nuova_riga.push('');
        nuova_riga.push(totScore);
        risultati.appendRow(nuova_riga);
      }
    }
  }
  
  quiz.submitGrades(formResponses);
  
  
}



