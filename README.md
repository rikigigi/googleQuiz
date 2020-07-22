# Usage

 - prepare a google spreadsheet with the following columns:

```
id question_text date_last errors_last notes ref answers				
```
 - prepare an empty template google quiz
 - prepare an empty template google document with the text `%TITOLOTEST%` , `%DESCRIZIONETEST%`, `%DATATEST%` in places that you like. They will be substituted with the title of the test, the description and the date.
 -  note the id of the spreadsheet and the templates (can be deduced from URL)
 - put in the `id` column an unique number for every question.
 - insert in the column `question_text` the question
 - insert the answers starting from the column `answers`
 - change the background of the correct answer with the rgb color #00FF00 (green)
 - change the brackground of the id of the questions that you want to use to make the quiz with the rgb color #FFE599 (some kind of yellow)
 - add a blank sheet to the spreadsheet and put it in the last position
 - open the file `quiz.js` with google apps script. Modify the `question_bank_ID` variable with the id of the spreadsheet that you noted before in the functions `prepare_new_quiz_from_selected` and `make_a_new_quiz`. Prepare a template with google documents, google forms and note their id. Put those id in the variables `quiz_template_ID` and `quiz_scritto_template_ID`
 - run the function `prepare_new_quiz_from_selected`
 - you will find that the last sheet now is written. You can put the title in the cell (1,1), the description in the cell(2,1) and the date in the cell (1,2)
   The numbers below are for custom marks of the tests. See source at [https://github.com/rikigigi/googleQuiz/blob/76dfd09af137719f43f2b1d380f5924f3852ed4a/quiz.js#L447]
 - run the function `make_a_new_quiz()`
 - enjoy it

The creation of the paper version and a draft of the presentation for the test correction is automatic.
