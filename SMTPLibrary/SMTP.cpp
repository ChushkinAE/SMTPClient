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

// ������ ������.
const string TEXT_EMPTY = "";
// ���������� SMTP-������.
const string SMTP_CLIENT_NAME = "ChushkinSMTP";
// ������������ ����.
const int PORT = 465;
// ������������ ����� �������� ������� ��� ��������� ������ �� ������� (� ��).
const int WAITING_TIME = 15000;

// ���������� SMTP-�������.
const string COMMAND_END = "\r\n";
// SMTP-������� �����������.
const string EHLO = "EHLO ";
// SMTP-������� ������ ��������������.
const string AUTH_LOGIN = "AUTH LOGIN ";
// SMTP-�������, ������������ �����������.
const string MAIL_FROM = "MAIL FROM:";
// SMTP-�������, ������������ ����������.
const string RCPT_TO = "RCPT TO:";
// SMTP-�������, ������������ ������ �������� ������ ������.
const string DATA = "DATA";
// ���������� ������, ������������ SMTP-��������.
const string DATA_END = COMMAND_END + "." + COMMAND_END;
// SMTP-�������, ����������� SMTP-�����.
const string QUIT = "QUIT";

// ��� ������ SMTP-�������, ��������������� ���������� ������� �������� � ��������.
const string CODE_SERVER_IS_REDY = "220";
// ��� ������ SMTP-�������, ��������������� "��������" ������ �� ���������� �������.
const string CODE_ACTION_OK = "250";
// ��� ������ SMTP-�������, ��������������� ���������� � ��������������.
const string CODE_AUTH_OK = "334";
// ��� ������ SMTP-�������, ��������������� ����������� ���������� ��������������.
const string CODE_AUTH_END_OK = "235";
// ��� ������ SMTP-�������, ��������������� ���������� ��������� ������ ������.
const string CODE_DATA_BEGIN = "354";
// ��� ������ SMTP-�������, ��������������� �������� ������ ��������.
const string CODE_CLOSE_CHANNEL = "221";

/**
 * ������������� SMTP-����������.
 */
void InitializeSMTP()
{
	SSL_library_init();
	ERR_load_BIO_strings();
	SSL_load_error_strings();
	OpenSSL_add_all_algorithms();
}

/**
 * �������� ������.
 * @param serverName - �������� ��� SMTP-�������.
 * @param serverIP - IP-����� SMTP-�������.
 * @param recipients - ���������� ������.
 * @param subject - ���� ������.
 * @param sender - ����� ����������� (�����).
 * @param password - ������ �����������.
 * @param message - ����� ������.
 * @param fileNames - ����� ������������ ������.
 * @return �������� ������, ���������� ��� �������� ������.
 * @note ��� ���������� �������� ������ ���������� ������ ������.
 * @note ������ ������������ ������������ "���������" ��� �� ��������� �����, ��� � �� IP-������. ���� �� ���� ���������� ������ ���� "����".
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

	// ����������� � SMTP-�������.
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

		return "�� ������� ������������ � �������. ";
	}
	else if (GetCode(answer) != CODE_SERVER_IS_REDY)
	{
		SendSMTPCommand(bio, QUIT + COMMAND_END);
		BIO_free_all(bio);
		SSL_CTX_free(ctx);

		return "������������ ����������� �������. ";
	}

	// ��������� ����������� SMTP-�������.
	string ehlo = EHLO + SMTP_CLIENT_NAME + COMMAND_END;
	bool res = ExecuteSMTPCommand(bio, ctx, ehlo, CODE_ACTION_OK);
	if (!res) return "�� ������� ���������������� ������. ";

	// ������ ��������������.
	res = ExecuteSMTPCommand(bio, ctx, AUTH_LOGIN + COMMAND_END, CODE_AUTH_OK);
	if (!res) return "�� ������� ������ ��������������. ";

	// �������� ������.
	string loginBase64 = base64_encode(sender, false);
	res = ExecuteSMTPCommand(bio, ctx, loginBase64 + COMMAND_END, CODE_AUTH_OK);
	if (!res) return "�� ������� ��������� �����. ";

	// �������� ������.
	string passwordBase64 = base64_encode(password, false);
	res = ExecuteSMTPCommand(bio, ctx, passwordBase64 + COMMAND_END, CODE_AUTH_END_OK);
	if (!res) return "�������������� ��������� ��� ��������� ��������� ������. ";

	// ����������� �����������.
	string from = MAIL_FROM + "<" + sender + ">" + COMMAND_END;
	res = ExecuteSMTPCommand(bio, ctx, from, CODE_ACTION_OK);
	if (!res) return "�� ������� ���������� �����������. ";

	// ����������� �����������.
	for (auto recipient : recipients)
	{
		string to = RCPT_TO + "<" + recipient + ">" + COMMAND_END;
		res = ExecuteSMTPCommand(bio, ctx, to, CODE_ACTION_OK);
		if (!res) return "�� ������� ���������� ���������� (" + recipient + ").";
	}

	// ����������� ������ �������� ������.
	res = ExecuteSMTPCommand(bio, ctx, DATA + COMMAND_END, CODE_DATA_BEGIN);
	if (!res) return "�� ������� ���������� ������ �������� ����������� ������. ";

	// �������� ������.
	res = SendData(bio, ctx, sender, subject, recipients, message, fileNames);
	if (!res) return "�� ������� ��������� ���������� ������. ";

	// ���������� ������.
	SendSMTPCommand(bio, QUIT + COMMAND_END);
	BIO_free_all(bio);
	SSL_CTX_free(ctx);

	return TEXT_EMPTY;
}

/**
 * ��������� ���������� ������� �������� ������ SendEmail.
 * @param serverName - �������� ��� SMTP-�������.
 * @param serverIP - IP-����� SMTP-�������.
 * @param recipients - ���������� ������.
 * @param subject - ���� ������.
 * @param sender - ����� ����������� (�����).
 * @param password - ������ �����������.
 * @param message - ����� ������.
 * @return �������� �������������� ������ �� ���������� �����������.
 * @note ��� ������������ ���� ���������� ������������ ������ ������.
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
		return "������� �������� ��� ��� IP-�����. ";
	else if (!serverName.empty() && !serverIP.empty())
		return "������ ������������ ������������ \"���������\" ��� �� ��������� �����, ��� � �� IP-������. \n�������� ���� �� �������� ����������. ";

	if (recipients.empty()) return "�������� ���� ������ ����������. ";
	if (subject.empty()) return "������� ���� ������. ";
	if (sender.empty()) return "������� ����� (�����). ";
	if (password.empty()) return "������� ������. ";
	if (message.empty()) return "������� ����� ������. ";

	if (subject.find(COMMAND_END) != -1) return "���� ������ �������� ����������� ���� �������� <CR><LF>. ";
	if (sender.find(COMMAND_END) != -1) return "����� �������� ����������� ���� �������� <CR><LF>. ";
	if (password.find(COMMAND_END) != -1) return "������ �������� ����������� ���� �������� <CR><LF>. ";
	for (auto recipient : recipients)
		if (recipient.find(COMMAND_END) != -1)
			return "���� �� ������� ����������� �������� ����������� ���� �������� <CR><LF> (" + recipient + "). ";

	return TEXT_EMPTY;
}

/**
 * �������� ������� SMTP-�������, ��������� ������ � ��� ���������.
 * @param bio - ���������� ������� �����/������.
 * @param ctx - ��������� ��� ��������� ����������� ����������.
 * @param command - �������.
 * @param code - ��������� ��� ������.
 * @return true - �����, false - ������.
 * @note � ������ ������� ��������� ���������� � ������������ ��� �������.
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
 * �������� ������� SMTP-�������.
 * @param bio - ���������� ������� �����/������.
 * @param command - �������, ������������ SMTP-�������.
 * @return ���������� �������������� ��������.
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
 * ��������� ������ SMTP-�������.
 * @param bio - ���������� ������� �����/������.
 * @return ����� SMTP-�������.
 * @note ������ ������ ������������ ������, ������� "����������" �� ���������� ����� �������� WAITING_TIME.
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
 * �������� ����������� ������.
 * @param bio - ���������� ������� �����/������.
 * @param ctx - ��������� ��� ��������� ����������� ����������.
 * @param sender - ����� �����������.
 * @param subject - ���� ������.
 * @param recipients - ���������� ������.
 * @param message - ����� ������.
 * @param fileNames - ����� ������������ ������.
 * @return true - �����, false - ������.
 * @note � ������ ������� ��������� ���������� � ������������ ��� �������.
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
 * �������� SMTP-������ �� �������������.
 * @param answer - ����� SMTP-�������.
 * @return true - ����� SMTP-������� ��������, false - ����� SMTP-������� ����������.
 * @note ��� ������������� ������ SMTP-������� ������������ false.
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
 * ��������� ���� ������ SMTP-�������.
 * @param answer - ����� SMTP-�������.
 * @return ��� ������ SMTP-������.
 * @note ��� ������������� ������ SMTP-������� ������������ ������ ������.
 */
string GetCode(const string &answer)
{
	if (answer.empty() || answer.length() < 3)
		return TEXT_EMPTY;

	return answer.substr(0, 3);
}