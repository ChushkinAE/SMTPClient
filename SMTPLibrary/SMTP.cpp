#include <Shlwapi.h>

#include <string>
#include <vector>
#include <fstream>

#include <openssl/ssl.h>
#include <openssl/bio.h>
#include <openssl/err.h>

#include <base64.h>

using namespace std;

string ValidateSendEmailArguments(
	const string& serverName,
	const string& serverIP,
	const vector<string>& recipients,
	const string& subject,
	const string& sender,
	const string& password,
	const string& message);
size_t SendSMTPCommand(BIO* bio, const string& command);
string GetSMTPAnswer(BIO* bio);
bool SendData(BIO* bio, SSL_CTX* ctx, const string& sender, const string& subject, const vector<string>& recipients, const string& message, const vector<string>& fileNames);
bool IsCommadComplete(const string& answer);
string GetCode(const string& answer);
bool ExecuteSMTPCommand(BIO* bio, SSL_CTX* ctx, const string& command, const string& code);

// Пустая строка.
const string TEXT_EMPTY = "";
// Именование SMTP-клента.
const string SMTP_CLIENT_NAME = "ChushkinSMTP";
// Используемый порт.
const int PORT = 465;
// Максимальное время отправки команды или получения ответа на команду (в мс).
const int WAITING_TIME = 15000;

// Завершение SMTP-команды.
const string COMMAND_END = "\r\n";
// SMTP-команда приветствия.
const string EHLO = "EHLO ";
// SMTP-команда начала аутентификации.
const string AUTH_LOGIN = "AUTH LOGIN ";
// SMTP-команда, обозначающая отправителя.
const string MAIL_FROM = "MAIL FROM:";
// SMTP-команда, обозначающая получателя.
const string RCPT_TO = "RCPT TO:";
// SMTP-команда, обозначающая начало передачи данных письма.
const string DATA = "DATA";
// Завершение данных, передаваемых SMTP-клиентом.
const string DATA_END = COMMAND_END + "." + COMMAND_END;
// SMTP-команда, завершающая SMTP-сеанс.
const string QUIT = "QUIT";

// Код ответа SMTP-сервера, соответствующий готовности сервера работать с клиентом.
const string CODE_SERVER_IS_REDY = "220";
// Код ответа SMTP-сервера, соответствующий "рядовому" ответу на корректную команду.
const string CODE_ACTION_OK = "250";
// Код ответа SMTP-сервера, соответствующий готовности к аутентификации.
const string CODE_AUTH_OK = "334";
// Код ответа SMTP-сервера, соответствующий корректному завершению аутентификации.
const string CODE_AUTH_END_OK = "235";
// Код ответа SMTP-сервера, соответствующий готовности принимать данные письма.
const string CODE_DATA_BEGIN = "354";
// Код ответа SMTP-сервера, соответствующий закрытию канала передачи.
const string CODE_CLOSE_CHANNEL = "221";

/**
 * Инициализация SMTP-библиотеки.
 */
void InitializeSMTP()
{
	SSL_library_init();
	ERR_load_BIO_strings();
	SSL_load_error_strings();
	OpenSSL_add_all_algorithms();
}

/**
 * Отправка письма.
 * @param serverName - доменное имя SMTP-сервера.
 * @param serverIP - IP-адрес SMTP-сервера.
 * @param recipients - получатели письма.
 * @param subject - тема письма.
 * @param sender - адрес отправителя (логин).
 * @param password - пароль отправителя.
 * @param message - текст письма.
 * @param fileNames - имена прикреплённых файлов.
 * @return Описание ошибки, полученной при отправки письма.
 * @note При корректной отправке письма возвращает пустую строку.
 * @note Нельзя одновременно использовать "адресацию" как по доменному имени, так и по IP-адресу. Один из этих аргументов должен быть "пуст".
 */
string SendEmail(
	const string& serverName,
	const string& serverIP,
	const vector<string>& recipients,
	const string& subject,
	const string& sender,
	const string& password,
	const string& message,
	const vector<string>& fileNames)
{
	string validateRes = ValidateSendEmailArguments(serverName, serverIP, recipients, subject, sender, password, message);
	if (!validateRes.empty()) return validateRes;

	// Подключение к SMTP-серверу.
	SSL_CTX* ctx = SSL_CTX_new(SSLv23_client_method());
	BIO* bio = BIO_new_ssl_connect(ctx);
	BIO_set_nbio(bio, 1);
	string connectionString = (serverIP.empty() ? serverName : serverIP) + ":" + to_string(PORT);
	BIO_set_conn_hostname(bio, connectionString.c_str());
	BIO_do_connect(bio);
	string answer = GetSMTPAnswer(bio);
	if (answer.empty())
	{
		BIO_free_all(bio);
		SSL_CTX_free(ctx);

		return "Не удалось подключиться к серверу. ";
	}
	else if (GetCode(answer) != CODE_SERVER_IS_REDY)
	{
		SendSMTPCommand(bio, QUIT + COMMAND_END);
		BIO_free_all(bio);
		SSL_CTX_free(ctx);

		return "Некорректное приветствие сервера. ";
	}

	// Получение приветствия SMTP-сервера.
	string ehlo = EHLO + SMTP_CLIENT_NAME + COMMAND_END;
	bool res = ExecuteSMTPCommand(bio, ctx, ehlo, CODE_ACTION_OK);
	if (!res) return "Не удалось поприветствовать сервер. ";

	// Начало аутентификации.
	res = ExecuteSMTPCommand(bio, ctx, AUTH_LOGIN + COMMAND_END, CODE_AUTH_OK);
	if (!res) return "Не удалось начать аутентификацию. ";

	// Отправка логина.
	string loginBase64 = base64_encode(sender, false);
	res = ExecuteSMTPCommand(bio, ctx, loginBase64 + COMMAND_END, CODE_AUTH_OK);
	if (!res) return "Не удалось отправить логин. ";

	// Отправка пароля.
	string passwordBase64 = base64_encode(password, false);
	res = ExecuteSMTPCommand(bio, ctx, passwordBase64 + COMMAND_END, CODE_AUTH_END_OK);
	if (!res) return "Аутентификация провалена или неудалось отправить пароль. ";

	// Обозначение отправителя.
	string from = MAIL_FROM + "<" + sender + ">" + COMMAND_END;
	res = ExecuteSMTPCommand(bio, ctx, from, CODE_ACTION_OK);
	if (!res) return "Не удалось обозначить отправителя. ";

	// Обозначение получателей.
	for (auto recipient : recipients)
	{
		string to = RCPT_TO + "<" + recipient + ">" + COMMAND_END;
		res = ExecuteSMTPCommand(bio, ctx, to, CODE_ACTION_OK);
		if (!res) return "Не удалось обозначить получателя (" + recipient + ").";
	}

	// Обозначение начала передачи данных.
	res = ExecuteSMTPCommand(bio, ctx, DATA + COMMAND_END, CODE_DATA_BEGIN);
	if (!res) return "Не удалось обозначить начало передачи содержимого письма. ";

	// Отправка данных.
	res = SendData(bio, ctx, sender, subject, recipients, message, fileNames);
	if (!res) return "Не удалось отправить содержимое письма. ";

	// Завершение сеанса.
	SendSMTPCommand(bio, QUIT + COMMAND_END);
	BIO_free_all(bio);
	SSL_CTX_free(ctx);

	return TEXT_EMPTY;
}

/**
 * Валидация аргументов функции отправки письма SendEmail.
 * @param serverName - доменное имя SMTP-сервера.
 * @param serverIP - IP-адрес SMTP-сервера.
 * @param recipients - получатели письма.
 * @param subject - тема письма.
 * @param sender - адрес отправителя (логин).
 * @param password - пароль отправителя.
 * @param message - текст письма.
 * @return Описание несоответствия одного из аргументов требованиям.
 * @note При корректности всех аргументов возвращается пустая строка.
 */
string ValidateSendEmailArguments(
	const string& serverName,
	const string& serverIP,
	const vector<string>& recipients,
	const string& subject,
	const string& sender,
	const string& password,
	const string& message)
{
	if (serverName.empty() && serverIP.empty())
		return "Введите доменное имя или IP-адрес. ";
	else if (!serverName.empty() && !serverIP.empty())
		return "Нельзя одновременно использовать \"адресацию\" как по доменному имени, так и по IP-адресу. \nОчистите один из эементов управления. ";

	if (recipients.empty()) return "Добавьте хоть одного получателя. ";
	if (subject.empty()) return "Введите тему письма. ";
	if (sender.empty()) return "Введите логин (почту). ";
	if (password.empty()) return "Введите пароль. ";
	if (message.empty()) return "Введите текст письма. ";

	if (subject.find(COMMAND_END) != -1) return "Тема письма содержит недоустимую пару символов <CR><LF>. ";
	if (sender.find(COMMAND_END) != -1) return "Логин содержит недоустимую пару символов <CR><LF>. ";
	if (password.find(COMMAND_END) != -1) return "Пароль содержит недоустимую пару символов <CR><LF>. ";
	for (auto recipient : recipients)
		if (recipient.find(COMMAND_END) != -1)
			return "Один из адресов пулочателей содержит недоустимую пару символов <CR><LF> (" + recipient + "). ";

	return TEXT_EMPTY;
}

/**
 * Отправка команды SMTP-серверу, получение ответа и его валидация.
 * @param bio - абстракция потоков ввода/вывода.
 * @param ctx - структура для установки защищённого соединения.
 * @param command - команда.
 * @param code - требуемый код ответа.
 * @return true - успех, false - провал.
 * @note В случае провала завершает соединение и высвобождает все ресурсы.
 */
bool ExecuteSMTPCommand(BIO* bio, SSL_CTX* ctx, const string &command, const string &code)
{
	size_t residue = SendSMTPCommand(bio, command);
	if (residue > 0)
	{
		BIO_free_all(bio);
		SSL_CTX_free(ctx);

		return false;
	}

	string answer = GetSMTPAnswer(bio);
	if (answer.empty() || GetCode(answer) != code)
	{
		BIO_free_all(bio);
		SSL_CTX_free(ctx);

		return false;
	}

	return true;
}

/**
 * Отправка команды SMTP-серверу.
 * @param bio - абстракция потоков ввода/вывода.
 * @param command - команда, отправляемая SMTP-серверу.
 * @return Количество неотправленных символов.
 */
size_t SendSMTPCommand(BIO* bio, const string& command)
{
	unsigned int startTime = clock();
	const char* buf = command.c_str();
	size_t size = command.length();
	while (size > 0)
	{
		unsigned int currentTime = clock();
		if ((currentTime - startTime) > WAITING_TIME)
			return size;

		int n = BIO_write(bio, buf, size);
		if (n > 0)
		{
			buf += n;
			size -= n;
		}
	}

	return 0;
}

/**
 * Получение ответа SMTP-сервера.
 * @param bio - абстракция потоков ввода/вывода.
 * @return Ответ SMTP-сервера.
 * @note Пустая строка соответстует ответу, который "перешагнул" за предельное время ожидания WAITING_TIME.
 */
string GetSMTPAnswer(BIO* bio)
{
	unsigned int startTime = clock();
	string res;
	char buf[256];
	int inputLength;
	do
	{
		unsigned int currentTime = clock();
		if ((currentTime - startTime) > WAITING_TIME)
			return TEXT_EMPTY;

		ZeroMemory(buf, sizeof(buf));
		inputLength = BIO_read(bio, buf, sizeof(buf) - 1);
		if (inputLength > 0)
		{
			buf[inputLength] = 0;
			string resPart(buf);
			res += resPart;
		}
	} while (!IsCommadComplete(res));

	return res;
}

/**
 * Отправка содержимого письма.
 * @param bio - абстракция потоков ввода/вывода.
 * @param ctx - структура для установки защищённого соединения.
 * @param sender - адрес отправителя.
 * @param subject - тема письма.
 * @param recipients - получатели письма.
 * @param message - текст письма.
 * @param fileNames - имена прикреплённых файлов.
 * @return true - успех, false - провал.
 * @note В случае провала завершает соединение и высвобождает все ресурсы.
 */
bool SendData(BIO* bio, SSL_CTX* ctx, const string &sender, const string &subject, const vector<string> &recipients, const string  &message, const vector<string> &fileNames)
{
	string boundary = "4324533534654776575656";
	string base = "MIME-Version: 1.0" + COMMAND_END + "Content-type: multipart/mixed; boundary=\"" + boundary +"\"" + COMMAND_END
		+ "Content-Transfer-Encoding: 8bit" + COMMAND_END;
	string from = "From: " + sender  + COMMAND_END;
	string tempSubject = "Subject: =?Windows-1251?B?" + base64_encode(subject) + "?=" + COMMAND_END;
	string to;
	for (const auto& recipient : recipients)
	{
		if (!to.empty())
		{
			to += ", ";
		}

		to += recipient;
	}

	to.insert(0, "To: ");
	to += COMMAND_END;
	string data = from + to + tempSubject + base + COMMAND_END + COMMAND_END;
	data += "--" + boundary + COMMAND_END  + "Content-Type: text/plain; charset=Windows-1251" + COMMAND_END
		+ "Content-Transfer-Encoding: 8bit" + COMMAND_END + COMMAND_END + message;
	size_t res = SendSMTPCommand(bio, data);
	if (res != 0)
	{
		BIO_free_all(bio);
		SSL_CTX_free(ctx);

		return false;
	}

	for (const auto& fileName : fileNames)
	{
		ifstream file(fileName, ios::binary);
		if (!file)
		{
			continue;
		}

		data = COMMAND_END + "--" + boundary + COMMAND_END + "Content-Type: application/octet-stream;" + COMMAND_END + "Content-Disposition: attachment; filename=\""
			+ PathFindFileName(fileName.c_str()) + "\"" + COMMAND_END + "Content-Transfer-Encoding: base64" + COMMAND_END + COMMAND_END;
		char buf[1024];
		string fileData;
		for (;;)
		{
			file.read(buf, size(buf));
			size_t size = file.gcount();
			if (size == 0)
				break;

			fileData.append(buf, size);
		}

		file.close();
		data += base64_encode(fileData, false) + COMMAND_END;
		res = SendSMTPCommand(bio, data);
		if (res != 0)
		{
			BIO_free_all(bio);
			SSL_CTX_free(ctx);

			return false;
		}
	}

	data = COMMAND_END + "--" + boundary +"--" + COMMAND_END + DATA_END;
	res = SendSMTPCommand(bio, data);
	if (res != 0)
	{
		BIO_free_all(bio);
		SSL_CTX_free(ctx);

		return false;
	}

	string answer = GetSMTPAnswer(bio);
	if (answer.empty() || GetCode(answer) != CODE_ACTION_OK)
	{
		BIO_free_all(bio);
		SSL_CTX_free(ctx);

		return false;
	}

	return true;
}

/**
 * Проверка SMTP-ответа на завершённость.
 * @param answer - ответ SMTP-сервера.
 * @return true - ответ SMTP-сервера завершён, false - ответ SMTP-сервера незавершён.
 * @note Для некорректного ответа SMTP-сервера возвращается false.
 */
bool IsCommadComplete(const string &answer)
{
	if (answer.empty() || answer.length() < 3)
		return false;

	if (answer[answer.length() - 2] != '\r' || answer[answer.length() - 1] != '\n')
		return false;

	if (answer[3] == ' ')
		return true;

	int lastCodeIndex;
	for (lastCodeIndex = answer.length() - 3; answer[lastCodeIndex - 1] != '\n' && answer[lastCodeIndex - 2] != '\r'
		&& lastCodeIndex > 1; --lastCodeIndex);

	if (answer[lastCodeIndex + 3] == ' ')
		return true;

	return false;
}

/**
 * Получение кода ответа SMTP-сервера.
 * @param answer - ответ SMTP-сервера.
 * @return Код ответа SMTP-сервер.
 * @note Для некорректного ответа SMTP-сервера возвращается пустая строка.
 */
string GetCode(const string &answer)
{
	if (answer.empty() || answer.length() < 3)
		return TEXT_EMPTY;

	return answer.substr(0, 3);
}