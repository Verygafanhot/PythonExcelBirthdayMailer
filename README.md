# PythonExcelBirthdayMailer
A simple python script to get Birthdays from an excel file and send an email every day at a chosen hour!

## Instalation

```bash
  pip install -r requirements.txt
```

## Configuration

On line 56-58 we will set up the email and password
<img width="830" alt="imagem" src="https://github.com/Verygafanhot/PythonExcelBirthdayMailer/assets/63465951/67727f7c-884d-4c07-bced-19aa8ee59fac">

On the sender write your email and for the password go here [HERE](https://myaccount.google.com/apppasswords)

Give it a nice name and click create

<img width="500" alt="imagem" src="https://github.com/Verygafanhot/PythonExcelBirthdayMailer/assets/63465951/08b3810f-92aa-42a9-973b-6e4e200bac95">

It will then give you a code!

Finally and most importantly open the excel file called birthday.xlsx

There you will find 3 example birthdays, you can get rid of them and use your own!

<img width="1434" alt="imagem" src="https://github.com/Verygafanhot/PythonExcelBirthdayMailer/assets/63465951/6004f285-5f79-47ee-b014-057c6ad8dce8">

### Don't remove the "END", move it to right after the last user on your list as shown in the included example


## Running

```bash
  python run.py
```
Now it will run until you stop it! Sending emails at the time you set it to (Default is 8AM)


## Personalization

### Personalize how the email looks
Make it look perfect! Just modify the HTML code found in lines 45-55

<img width="434" alt="imagem" src="https://github.com/Verygafanhot/PythonExcelBirthdayMailer/assets/63465951/0b800a15-827b-45d9-bab1-efc22f81ea26">

### Change the hour of execution
Don't like 8AM?! Just change it on line 88!

<img width="342" alt="imagem" src="https://github.com/Verygafanhot/PythonExcelBirthdayMailer/assets/63465951/f4212d30-bf6f-428c-b4bc-fa2ef1548401">

