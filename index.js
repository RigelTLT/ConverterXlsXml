const puppeteer = require("puppeteer");
const fs = require("fs");
const xlsx = require("xlsx");
const xml2js = require("xml2js");
const axios = require("axios");
require("dotenv").config();

// URL сайта
const siteUrl = "https://hb.p-cod.com:18181/clog2/";

// Шаг 1: Авторизация и клик на кнопку
async function downloadXLS() {
  const browser = await puppeteer.launch({ headless: false }); // Запускаем браузер в режиме не в фоне для отладки
  const page = await browser.newPage();

  // Открываем страницу и авторизуемся
  await page.goto(siteUrl);

  // Заполняем форму авторизации
  await page.type("#Login", process.env.Login); // Укажите ваш логин
  await page.type("#Pass", process.env.Password); // Укажите ваш пароль
  await page.click(".loginButton"); // Замените на соответствующий селектор для кнопки авторизации

  // Ждем загрузки страницы после авторизации
  await page.waitForNavigation();

  console.log("Авторизация прошла успешно!");

  await page.click("#mainpage:j_id75");
  await page.waitForTimeout(5000); // Ждем 5 секунд, чтобы файл успел загрузиться
  // Шаг 2: Нажимаем на кнопку для выгрузки Excel
  await page.click("#containersForm\\:containersTable\\:j_id340"); // Используем экранирование для двойных двоеточий в селекторах

  // Ждем скачивания файла (можно задать ожидание на определенную загрузку или файл)
  await page.waitForTimeout(5000); // Ждем 5 секунд, чтобы файл успел загрузиться

  console.log("Файл выгружен в Excel");

  // Скачиваем файл из папки с загрузками (обычно это папка по умолчанию)
  // Если файл сохраняется на диске, то указываем путь и сохраняем его локально
  const filePath = "/path/to/downloaded/file.xlsx"; // Замените на реальный путь

  // Сохраняем файл на диск (если доступ к файлу)
  fs.writeFileSync("data.xls", filePath);

  // Закрываем браузер
  await browser.close();
}

// Шаг 3: Конвертация XLS в XML
function convertXLSToXML() {
  const workbook = xlsx.readFile("data.xls");
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(worksheet);

  const builder = new xml2js.Builder();
  const xmlData = builder.buildObject({ Root: { Entry: jsonData } });

  fs.writeFileSync("data.xml", xmlData);
  console.log("Файл XML сохранен как data.xml");
}

// Шаг 4: Отправка XML на другой сервис
async function sendXMLToService() {
  const xmlData = fs.readFileSync("data.xml", "utf8");
  const apiUrl = "https://example.com/api/endpoint"; // Замените на URL сервиса

  try {
    const response = await axios.post(apiUrl, xmlData, {
      headers: {
        "Content-Type": "application/xml",
      },
    });
    console.log("XML успешно отправлен на сервис");
  } catch (error) {
    console.error("Ошибка при отправке XML:", error);
    throw error;
  }
}

// Основная функция, которая объединяет все шаги
async function main() {
  try {
    // Шаг 1: Авторизация и скачивание файла
    await downloadXLS();

    // Шаг 2: Конвертация XLS в XML
    convertXLSToXML();

    // Шаг 3: Отправка XML на другой сервис
    await sendXMLToService();
  } catch (error) {
    console.error("Произошла ошибка при обработке:", error);
  }
}

// Запуск основной функции
main();
