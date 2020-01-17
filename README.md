# Usage

 - prepare a google spreadsheet with the following columns:

```
id question_text date_last errors_last notes ref answers				
```
   note the id of the spreadsheet (can be deduced from URL)

 - put in the `id` column an unique number for every question.
 - insert in the column `question_text` the question
 - insert the answers starting from the column `answers`
 - change the background of the correct answer with the rgb color #00FF00 (green)
 - select answer that you want to use to make the quiz with the rgb color #FFE599 (some kind of yellow)
 - add a blank sheet to the spreadsheet and put it in the last position
 - open the file `quiz.js` with google apps script. Modify the `question_bank_ID` variable with the id of the spreadsheet that you noted before in the functions `prepare_new_quiz_from_selected` and `make_a_new_quiz`. Prepare a template with google documents, google forms and note their id. Put those id in the variables `quiz_template_ID` and `quiz_scritto_template_ID`
 - run the function `prepare_new_quiz_from_selected`
 - run the function `make_a_new_quiz()`
 - enjoy it
