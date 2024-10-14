<p align="center">
  <img src="img.shields.io/image/javaScript_logo.png" width="200" alt="Node Logo" /></a>
</p>

<p align="center">
		<em>Developed with the software and tools below.</em>
</p>

<p align="center">
    <img src="img.shields.io/badge/Google.svg?style=flat&logo=vitest&logoColor=white" alt="Google">
    <img src="img.shields.io/badge/GoogleSheets.svg?style=flat&logo=vitest&logoColor=white" alt="Google Sheets">
    <img src="img.shields.io/badge/GoogleCalendar.svg?style=flat&logo=vitest&logoColor=white" alt="Google Calendar">
</p>

## 📝 Overview

[EN]

This code is very simple and was developed due to the need of the events department of my current company to capture data from a spreadsheet (Google Sheets) and transform it into an event in Google Calendar.

With the new update, the code allows you to read data from different tabs and save it in a specific category in Google Calendar.

NOTE: For the code to work, you need to have the Google Workspace environment and always send the code with the .gs extension and have created the calendars you want within Google Calendar.

[PT-BR]

Esse código é muito simples e foi desenvolvido através da necessidade do setor de eventos da minha atual empresa de capturar os dados de uma planilha (Google Sheets) e transformar em um evento no Google Calendar.

Com a nova atualização, o codigo permite ler dados de abas diferentes e salvar em uma categoria especifica no Google Calendar.

OBS: Para o codigo funcionar, voce precisa ter o ambiente Google Workspace e sempre enviar o cogido com a extenção .gs e ter criado as agendas que deseja dentro do Google Calendar.

---

## 📦️ Concept and application

Within the Google Workspace environment, it is possible to integrate your tools through manual processes or automatically through scripts. Within Google Sheets, there is an option to register scripts to execute a task based on functions.

- Fetch values ​​from the entire spreadsheet (One or more tabs)
- Call Google Calendar
- Read the title of all columns
- Converts the time and duration of the event
- Checks if the event already exists to avoid it
- Prohibit the creation of an event on a past date
- Set a color for the event based on the status
- Save the event in a specific category
- Creates the event with the details and sends the email to the user responsible for the registered event
- Reading and notification by email when there is an update in the spreadsheet
---

## 📖 Challenge

Google's API is very vast and many functions used in the code are derived from Google itself, so I needed to be patient to understand and study how each of them works. Certain functions, as I will list below, are new to me.

Just a note, Google's APPS Script tool already has a debugging option, so you don't need to run tests and the spreadsheet needs to be open. You can also create a button to execute the function and schedule it to run at a certain time.

```sh
getActiveSpreadsheet()
getActiveSheet()
getDataRange()
getValues()
Logger.log()
```
---

## 🎉 Acknowledgements

I would like to thank the entire events department at my company for providing me with this challenge by giving me what I needed and trusting me to deliver.

More automations like these will soon be available at the company and I am super excited to learn more about this Google environment.
