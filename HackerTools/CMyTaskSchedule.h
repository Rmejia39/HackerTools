#pragma once
#include <Atlbase.h>
#include <comdef.h>
#include <taskschd.h>
#pragma comment(lib, "taskschd.lib")
class CMyTaskSchedule
{
public:
	CMyTaskSchedule();
	~CMyTaskSchedule();
	//�����ƻ�����
	BOOL NewTask(char* lpszTaskName, char* lpszProgramPath, char* lpszParameters, char* lpszAuthor);

	//ɾ���ƻ�����
	BOOL Delete(char* lpszTaskName);

	//�ж�ָ���ƻ������Ƿ���
	BOOL IsTaskValid(char *lpszTaskName);

	// ����ָ������ƻ�
	BOOL Run(char *lpszTaskName, char *lpszParam);

	// �ж�ָ������ƻ��Ƿ�����
	BOOL IsEnable(char *lpszTaskName);

	// ����ָ������ƻ��Ƿ��������ǽ���
	BOOL SetEnable(char *lpszTaskName, BOOL bEnable);

private:
	ITaskService *m_lpITS;
	ITaskFolder *m_lpRootFolder;
};

