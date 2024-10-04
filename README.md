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

[PT-BR]

Esse código é muito simples e foi desenvolvido através da necessidade do setor de eventos da minha atual empresa de capturar os dados de uma planilha (Google Sheets) e transformar em um evento no Google Calendar.

---

## 📦️ Concept and application

Within the Google Workspace environment, it is possible to integrate your tools through manual processes or automatically through scripts. Within Google Sheets, there is an option to register scripts to execute a task based on functions.

- Searches for values ​​from the entire spreadsheet
- Calls Google Calendar
- Reads the title of all columns
- Converting event time and duration
- Checks if the event already exists to avoid to avoid
- Sets a color for the event based on the status
- Create the event with the details and send the email to the user responsible for the registered event
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