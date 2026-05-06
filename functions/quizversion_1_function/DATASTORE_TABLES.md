# Quiz Function – Datastore Tables

In **Zoho Catalyst**, use your project’s **Datastore** and create these **tables** (exact names matter).

---

## 1. **CreateQuiz** (required for create/join/load quiz)

Used to store quiz metadata and questions. **Required** for creating quizzes and for “Join Game” to find a quiz by code.

| Column         | Type   | Used for |
|----------------|--------|----------|
| QuizTitle      | String | Quiz title |
| Description    | String | Quiz description |
| Question       | String | Question text |
| option1        | String | Option 1 |
| option2        | String | Option 2 |
| option3        | String | Option 3 |
| option4        | String | Option 4 |
| CorrectAnswer  | String | Correct option index (e.g. "0", "1") |
| AddQuestion    | String | Quiz ID (links rows of same quiz) |
| StartQuiz      | String | **Game code** (e.g. 82Z0OT) – used to find quiz on join |
| Timelimit      | String | Time limit in seconds (e.g. "30") |
| ImageFileId    | String | Optional image file ID |
| MODIFIEDTIME   | DateTime | Auto or manual |

**Note:** One row per question. Same quiz = same `AddQuestion` and `StartQuiz`.

---

## 2. **Question** (optional)

Extra table for structured question data. If you don’t create it, quiz still works; inserts here are best-effort.

| Column       | Type   |
|--------------|--------|
| QuizId       | String |
| GameCode     | String |
| QuestionId   | String |
| Question     | String |
| option1–option4 | String |
| CorrectAnswer| String |
| Points       | String |
| Timelimit    | String |
| QuestionOrder| String |
| ImageUrl     | String |
| MODIFIEDTIME | DateTime |

---

## 3. **QuizResults** (optional – for saving results when quiz ends)

| Column          | Type   |
|-----------------|--------|
| QuizId          | String |
| GameCode        | String |
| QuizTitle       | String |
| PlayerId        | String |
| PlayerName      | String |
| Avatar          | String |
| Ranking         | String |
| Score           | String |
| CorrectAnswers  | String |
| TotalQuestions  | String |
| QuestionResults| String |
| CompletedDate   | String |
| MODIFIEDTIME    | DateTime |

---

## 4. **PlayerName** (optional – for saving each answer)

| Column        | Type   |
|---------------|--------|
| Name          | String |
| GameCode      | String |
| avatar        | String |
| Question      | String |
| Correctanswer | String |
| score         | String |
| MODIFIEDTIME  | DateTime |

---

## 5. **quiz_feedback** (optional – for feedback after quiz)

| Column     | Type   |
|------------|--------|
| quizCode   | String |
| playerId   | String |
| playerName | String |
| rating     | String |
| comment    | String |

---

## Minimum to get “Join Game” working

- Create the **CreateQuiz** table with the columns above (at least: QuizTitle, Description, Question, option1–option4, CorrectAnswer, AddQuestion, StartQuiz, Timelimit; MODIFIEDTIME if required by Catalyst).

The other tables are optional for extra features (results, per-answer log, feedback).
