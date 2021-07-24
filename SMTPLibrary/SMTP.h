#pragma once

#ifndef __SMTP__
#define __SMTP__

#include <string>
#include <vector>

using namespace std;

/**
 * ������������� SMTP-����������.
 */
void InitializeSMTP();

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
	const vector<string>& fileNames);

#endif /* __SMTP__ */