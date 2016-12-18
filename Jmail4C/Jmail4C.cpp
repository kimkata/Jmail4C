// Jmail4C.cpp : 定义控制台应用程序的入口点。
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
#include <comutil.h>	//在Jmail日志文件输出时需要用此头文件
#import "jmail.dll"

using namespace std;
using namespace boost;
using namespace boost::xpressive;

//发送邮件的方法，注意：这里没有public，如果有public，则必须在头文件中添加声明！
//邮箱地址的正则表达式，从网站获取的。
string reg_for_mailaddress = "[\\w!#$%&'*+/=?^_`{|}~-]+(?:\\.[\\w!#$%&'*+/=?^_`{|}~-]+)*@(?:[\\w](?:[\\w-]*[\\w])?\\.)+[\\w](?:[\\w-]*[\\w])?";
sregex reg = sregex::compile(reg_for_mailaddress);
//发送邮件的结构体函数
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
		return "错误的邮箱地址";
	}

	{
		jmail::IMessagePtr  pMessage("Jmail.Message");
		try {
			pMessage->From = this->sender;
			pMessage->FromName = this->sendername;

			//检查邮件的附件是否存在,只有当附件存在，才发送邮件，否则不发送附件
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
			//如果没有发送成功，则返回成功；否则输出失败，并在console窗口中输出发送的日志，以便查看失败的原因。
			if (pMessage->Send(this->smtp, VARIANT_FALSE)) {
				pMessage->Close();
				return "成功√";
			}
			else {
				pMessage->Close();
				return "失败×";
				cout << _com_util::ConvertBSTRToString((pMessage->Log).Detach());

			}
		}
		//异常处理模块，输出错误的信息
		catch (_com_error & E)
		{
			cerr
				<< "Error: 0x" << hex << E.Error() << endl
				<< E.ErrorMessage() << endl;
			cout << _com_util::ConvertBSTRToString((pMessage->Log).Detach());
			pMessage->Close();
			return "->发生异常";
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

	//如果setup4mail存在，则读取前4行，存入到buff[]中
	if (setup4mail) {
		std::cout << "------------------用户名+密码配置信息如下：-------------------------" << endl;
		for (int i = 0; i < 4; i++) {
			getline(setup4mail, buff[i], '\n');
			std::cout << buff[i] << endl;
		}
		std::cout << "---------------------------------------------------------------------" << endl;
	}
	if (regex_search(buff[1], reg)) {
		sregex reg_username_head = sregex::compile("用户名:");
		sregex reg_mailaddress_head = sregex::compile("邮箱地址:");
		sregex reg_password_head = sregex::compile("登陆密码:");
		sregex reg_smtp_head = sregex::compile("邮件服务器:");

		username = regex_replace(buff[0], reg_username_head, "");
		mailAddress = regex_replace(buff[1], reg_mailaddress_head, "");
		password = regex_replace(buff[2], reg_password_head, "");
		smtp = regex_replace(buff[3], reg_smtp_head, "");

		std::cout << "//////////////////////////////////////////////////////////////////////////////" << endl;
		std::cout << "用户名：" << username << endl;
		std::cout << "邮箱地址：" << mailAddress << endl;
		std::cout << "邮箱密码：" << password << endl;
		std::cout << "邮件SMTP服务器：" << smtp << endl;
		std::cout << "------------------请核对以上信息，如果都正确，请按任意键继续--------------------" << endl;
	}
	if (maildata) {
		std::cout << "正在统计邮件的数量：" << endl;

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
		std::cout << endl << "邮件列表的标题行为：" << endl
			<< "--------------------------------------------------------------------------------------------" << endl
			<< first_line_without_tab << endl << "--------------------------------------------------------------------------------------------" << endl
			<< "备注：默认第一行为标题行，以上标题行不统计在邮件数量中！" << endl << endl
			<< "邮件的数量为:" << vector_of_line.size()-1 << endl << endl
			<< "请检查最后一行是否为空行，如果是空行，请将其删除！" << endl << endl;
	}



	cout << "如果信息无误，按任意键开始发送邮件：" << endl << endl;
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
		//每一封邮件变化的只有以下几个部分
		MyJmail.recepient_name = (*citer)[0].c_str();
		MyJmail.recepient = (*citer)[1].c_str();
		MyJmail.subject = (*citer)[2].c_str();
		//第三项是文件名，可以不使用
		MyJmail.body = (*citer)[4].c_str();
		MyJmail.attachment = (*citer)[5].c_str();
		vector_of_mail_struct.push_back(MyJmail);
	}
	std::cout << "已经获取邮件列表，准备发送邮件！" << endl << endl;
	log4mail.open("log4mail.txt", ios::out | ios::trunc);

	int i = 0;
	std::cout << "邮件发送进度条：";
	progress_display pd(vector_of_mail_struct.size());
	for (vector<mymail>::iterator citer = vector_of_mail_struct.begin(); citer != vector_of_mail_struct.end(); citer++)
	{
		string result = citer->send();	
		i++;		
		++pd;
		//std::cout << citer->recepient_name << ":	" << citer->recepient << "	\n发送状态->" << result << "	计数：" << i << endl;
		//std::system("cls");
		log4mail << citer->recepient_name << ":	" << citer->recepient << "	发送状态->" << result << "	计数：" << i << endl;
	}
	log4mail.close();
	std::cout << endl << endl << "程序执行完毕！" << endl << endl;
	std::system("pause");
	CoUninitialize();
	return 0;
}


