#pragma once

#ifndef __SMTP__
#define __SMTP__

#include <string>
#include <vector>

using namespace std;

/**
 * Инициализация SMTP-библиотеки.
 */
void InitializeSMTP();

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
	const vector<string>& fileNames);

#endif /* __SMTP__ */