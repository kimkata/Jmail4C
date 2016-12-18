// Jmail4C.cpp : �������̨Ӧ�ó������ڵ㡣
//

#include "stdafx.h"
#include <iostream>
#include <fstream>
#include <string>
#include <regex>
#include <vector>
#include <boost\xpressive\xpressive_dynamic.hpp>
#include <boost\tokenizer.hpp>
#include <boost\typeof\typeof.hpp>
#include <boost\filesystem.hpp>
#include <boost\progress.hpp>
#include <comutil.h>	//��Jmail��־�ļ����ʱ��Ҫ�ô�ͷ�ļ�
#import "jmail.dll"

using namespace std;
using namespace boost;
using namespace boost::xpressive;

//�����ʼ��ķ�����ע�⣺����û��public�������public���������ͷ�ļ������������
//�����ַ��������ʽ������վ��ȡ�ġ�
string reg_for_mailaddress = "[\\w!#$%&'*+/=?^_`{|}~-]+(?:\\.[\\w!#$%&'*+/=?^_`{|}~-]+)*@(?:[\\w](?:[\\w-]*[\\w])?\\.)+[\\w](?:[\\w-]*[\\w])?";
sregex reg = sregex::compile(reg_for_mailaddress);
//�����ʼ��Ľṹ�庯��
struct  mymail {
	const char *sender;
	const char *sendername;
	const char *mailserverusername;
	const char *mailserverpassword;
	const char *smtp;
	const char *recepient;
	const char *recepient_name;
	const char *subject;
	const char *body;
	const char *attachment;
	string send();

};
string mymail::send()
{
	string s = this->recepient;	
	if (!regex_match(s, reg))
	{
		return "����������ַ";
	}

	{
		jmail::IMessagePtr  pMessage("Jmail.Message");
		try {
			pMessage->From = this->sender;
			pMessage->FromName = this->sendername;

			//����ʼ��ĸ����Ƿ����,ֻ�е��������ڣ��ŷ����ʼ������򲻷��͸���
			if (boost::filesystem::exists(this->attachment)) {
				pMessage->AddAttachment(this->attachment, VARIANT_FALSE, "");				

			}

			pMessage->Subject = this->subject;
			pMessage->Body = this->body;
			pMessage->AddRecipient(this->recepient, this->recepient_name, "");
			pMessage->MailServerUserName = this->mailserverusername;
			pMessage->MailServerPassWord = this->mailserverpassword;
			pMessage->Priority = 3;
			pMessage->Charset = "GB2312";
			pMessage->Encoding = "base64";
			pMessage->Logging = true;
			//���û�з��ͳɹ����򷵻سɹ����������ʧ�ܣ�����console������������͵���־���Ա�鿴ʧ�ܵ�ԭ��
			if (pMessage->Send(this->smtp, VARIANT_FALSE)) {
				pMessage->Close();
				return "�ɹ���";
			}
			else {
				pMessage->Close();
				return "ʧ�ܡ�";
				cout << _com_util::ConvertBSTRToString((pMessage->Log).Detach());

			}
		}
		//�쳣����ģ�飬����������Ϣ
		catch (_com_error & E)
		{
			cerr
				<< "Error: 0x" << hex << E.Error() << endl
				<< E.ErrorMessage() << endl;
			cout << _com_util::ConvertBSTRToString((pMessage->Log).Detach());
			pMessage->Close();
			return "->�����쳣";
		}
	}
}


int main(int argc, char* *argv)
{
	string buff[4];
	string mailAddress;
	string username;
	string password;
	string smtp;
	ifstream setup4mail("setup4mail.txt");
	ifstream maildata("maildata.txt");
	ofstream log4mail;
	string str;
	vector<string> vector_of_line;
	vector<mymail> vector_of_mail_struct;
	CoInitialize(NULL);

	//���setup4mail���ڣ����ȡǰ4�У����뵽buff[]��
	if (setup4mail) {
		std::cout << "------------------�û���+����������Ϣ���£�-------------------------" << endl;
		for (int i = 0; i < 4; i++) {
			getline(setup4mail, buff[i], '\n');
			std::cout << buff[i] << endl;
		}
		std::cout << "---------------------------------------------------------------------" << endl;
	}
	if (regex_search(buff[1], reg)) {
		sregex reg_username_head = sregex::compile("�û���:");
		sregex reg_mailaddress_head = sregex::compile("�����ַ:");
		sregex reg_password_head = sregex::compile("��½����:");
		sregex reg_smtp_head = sregex::compile("�ʼ�������:");

		username = regex_replace(buff[0], reg_username_head, "");
		mailAddress = regex_replace(buff[1], reg_mailaddress_head, "");
		password = regex_replace(buff[2], reg_password_head, "");
		smtp = regex_replace(buff[3], reg_smtp_head, "");

		std::cout << "//////////////////////////////////////////////////////////////////////////////" << endl;
		std::cout << "�û�����" << username << endl;
		std::cout << "�����ַ��" << mailAddress << endl;
		std::cout << "�������룺" << password << endl;
		std::cout << "�ʼ�SMTP��������" << smtp << endl;
		std::cout << "------------------��˶�������Ϣ���������ȷ���밴���������--------------------" << endl;
	}
	if (maildata) {
		std::cout << "����ͳ���ʼ���������" << endl;

		while (!maildata.eof())
		{

			getline(maildata, str, '\n');
			if (!(str == ""))
			{
				vector_of_line.push_back(str);
				std::cout << "*";
			}

		}
		sregex tab = sregex::compile("\t");
		string first_line_without_tab = regex_replace(vector_of_line[0], tab, "  |  ");
		std::cout << endl << "�ʼ��б�ı�����Ϊ��" << endl
			<< "--------------------------------------------------------------------------------------------" << endl
			<< first_line_without_tab << endl << "--------------------------------------------------------------------------------------------" << endl
			<< "��ע��Ĭ�ϵ�һ��Ϊ�����У����ϱ����в�ͳ�����ʼ������У�" << endl << endl
			<< "�ʼ�������Ϊ:" << vector_of_line.size()-1 << endl << endl
			<< "�������һ���Ƿ�Ϊ���У�����ǿ��У��뽫��ɾ����" << endl << endl;
	}



	cout << "�����Ϣ���󣬰��������ʼ�����ʼ���" << endl << endl;
	std::system("pause");
	std::system("cls");


	struct mymail MyJmail;
	vector<vector<string>> myMailData;
	char_separator<char> sep("\t");


	for (vector<string>::const_iterator citer = vector_of_line.begin(); citer!= vector_of_line.end();citer++) {
		vector<string> mailstr;
		tokenizer<char_separator<char>> tok(*citer, sep);
		for (BOOST_AUTO(pos, tok.begin()); pos != tok.end(); ++pos) {
			string s = *pos;
			mailstr.push_back(s);
		}
		myMailData.push_back(mailstr);
	}


	for (vector<vector<string>>::const_iterator citer = myMailData.begin(); citer != myMailData.end(); citer++)
	{
		MyJmail.sender = mailAddress.c_str();
		MyJmail.sendername = username.c_str();
		MyJmail.mailserverusername = username.c_str();
		MyJmail.mailserverpassword = password.c_str();
		MyJmail.smtp = smtp.c_str();
		//ÿһ���ʼ��仯��ֻ�����¼�������
		MyJmail.recepient_name = (*citer)[0].c_str();
		MyJmail.recepient = (*citer)[1].c_str();
		MyJmail.subject = (*citer)[2].c_str();
		//���������ļ��������Բ�ʹ��
		MyJmail.body = (*citer)[4].c_str();
		MyJmail.attachment = (*citer)[5].c_str();
		vector_of_mail_struct.push_back(MyJmail);
	}
	std::cout << "�Ѿ���ȡ�ʼ��б�׼�������ʼ���" << endl << endl;
	log4mail.open("log4mail.txt", ios::out | ios::trunc);

	int i = 0;
	std::cout << "�ʼ����ͽ�������";
	progress_display pd(vector_of_mail_struct.size());
	for (vector<mymail>::iterator citer = vector_of_mail_struct.begin(); citer != vector_of_mail_struct.end(); citer++)
	{
		string result = citer->send();	
		i++;		
		++pd;
		//std::cout << citer->recepient_name << ":	" << citer->recepient << "	\n����״̬->" << result << "	������" << i << endl;
		//std::system("cls");
		log4mail << citer->recepient_name << ":	" << citer->recepient << "	����״̬->" << result << "	������" << i << endl;
	}
	log4mail.close();
	std::cout << endl << endl << "����ִ����ϣ�" << endl << endl;
	std::system("pause");
	CoUninitialize();
	return 0;
}


