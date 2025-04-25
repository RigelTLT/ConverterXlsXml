const axios = require("axios");
const cheerio = require("cheerio");
const xlsx = require("xlsx");
const xml2js = require("xml2js");
const fs = require("fs");

// Создаем сессию для сохранения cookies
const session = axios.create({
  withCredentials: true, // для отправки cookies
});

// Шаг 1: Авторизация
async function login(username, password) {
  const loginUrl = "https://hb.p-cod.com:18181/clog2/j_security_check";

  const loginData = {
    j_username: username,
    j_password: password,
  };

  try {
    const response = await session.post(
      loginUrl,
      new URLSearchParams(loginData)
    );
    console.log("Авторизация успешна!");
    return response;
  } catch (error) {
    console.error("Ошибка при авторизации:", error);
    throw error;
  }
}

// Шаг 2: Нажатие на кнопку для выгрузки данных
async function fetchXLSData() {
  const formUrl = "https://hb.p-cod.com:18181/clog2/containersForm";
  const formData = {
    "containersForm:containersTable:j_id735": "submit",
  };

  try {
    const response = await session.post(formUrl, new URLSearchParams(formData));
    console.log("Данные выгружены успешно!");
    return response;
  } catch (error) {
    console.error("Ошибка при выгрузке данных:", error);
    throw error;
  }
}

// Шаг 3: Скачивание XLS-файла
async function downloadXLS() {
  const xlsUrl = "https://hb.p-cod.com:18181/clog2/path_to_xls_file"; // Замените на актуальный путь

  try {
    const response = await session.get(xlsUrl, { responseType: "arraybuffer" });
    fs.writeFileSync("data.xls", response.data);
    console.log("Файл XLS сохранен как data.xls");
  } catch (error) {
    console.error("Ошибка при скачивании XLS-файла:", error);
    throw error;
  }
}

// Шаг 4: Конвертация XLS в XML
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

// Шаг 5: Отправка XML на другой сервис
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
    const username = "your_username";
    const password = "your_password";

    // Шаг 1: Авторизация
    await login(username, password);

    // Шаг 2: Нажатие на кнопку для выгрузки данных
    await fetchXLSData();

    // Шаг 3: Скачивание XLS
    await downloadXLS();

    // Шаг 4: Конвертация XLS в XML
    convertXLSToXML();

    // Шаг 5: Отправка XML на другой сервис
    await sendXMLToService();
  } catch (error) {
    console.error("Произошла ошибка при обработке:", error);
  }
}

// Запуск основной функции
main();
