const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const xml2js = require("xml2js");
const axios = require("axios");
require("dotenv").config();

// Папка загрузок
const downloadsFolder = path.join("C:", "Users", "DK", "Downloads"); // Путь к папке с загрузками

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
  await page.waitForSelector("#mainpage\\:j_id75", {
    timeout: 5000,
  });
  // Ожидаем появления кнопки для выгрузки
  await page.click("#mainpage\\:j_id75");
  await page.waitForSelector("#containersForm\\:containersTable\\:j_id340", {
    timeout: 5000,
  }); // Ждем 5 секунд, чтобы элемент стал доступным

  // Шаг 2: Ожидаем, что при клике откроется новое окно
  await page.waitForFunction(() => {
    const button = document.querySelector(
      "#containersForm\\:containersTable\\:j_id340"
    );
    if (button) {
      button.click();
      return true; // Успешное выполнение клика
    }
    return false;
  });

  console.log("Нажата кнопка для выгрузки файла.");

  // Ждем, пока файл будет доступен
  try {
    await page.waitForSelector('a[href*="ContainersToExcel"]', {
      timeout: 20000,
    });
  } catch {
    console.error("Ожидание скачивания");
  }

  console.log("Файл должен быть выгружен в Excel");

  // Ожидаем, пока файл появится в папке загрузок
  const filePath = await findDownloadedFile();

  if (filePath) {
    console.log(`Файл найден: ${filePath}`);
  } else {
    console.error("Файл не найден в папке загрузок");
  }

  // Закрываем браузер
  await browser.close();
}

// Функция для поиска файла в папке загрузок
function findDownloadedFile() {
  return new Promise((resolve, reject) => {
    fs.readdir(downloadsFolder, (err, files) => {
      if (err) {
        reject(err);
        return;
      }

      const file = files.find(
        (f) => f.startsWith("ClientContainers") && f.endsWith(".xls")
      );
      if (file) {
        resolve(path.join(downloadsFolder, file)); // Возвращаем полный путь к файлу
      } else {
        resolve(null); // Файл не найден
      }
    });
  });
}

// Шаг 3: Конвертация XLS в XML
function convertXLSToXML(filePath) {
  const workbook = xlsx.readFile(filePath);
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
    const filePath = await findDownloadedFile();
    if (filePath) {
      convertXLSToXML(filePath);

      // Шаг 3: Отправка XML на другой сервис
      await sendXMLToService();
    } else {
      console.error("Не удалось найти файл для конвертации");
    }
  } catch (error) {
    console.error("Произошла ошибка при обработке:", error);
  }
}

// Запуск основной функции
main();
