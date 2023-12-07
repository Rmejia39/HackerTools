#include "stdafx.h"
#include "CMyTaskSchedule.h"



//************************************************************
// ��������: CMyTaskSchedule 
// ����˵��: ��ʼ���ƻ�����
// ��	 ��: GuiShou
// ʱ	 ��: 2019/1/22
// ��	 ��: void
// �� �� ֵ: void
//************************************************************
CMyTaskSchedule::CMyTaskSchedule()
{
	m_lpITS = NULL;
	m_lpRootFolder = NULL;
	//��ʼ��COM�ӿڻ���
	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		MessageBoxA(NULL,"CoInitialize","Error",MB_OK);
	}

	//��������������
	hr = CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER,
		IID_ITaskService, (LPVOID*)(&m_lpITS));
	if (FAILED(hr))
	{
		MessageBoxA(NULL, "CoCreateInstance", "Error", MB_OK);
	}

	//���ӵ����������
	hr = m_lpITS->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
	if (FAILED(hr))
	{
		MessageBoxA(NULL, "Connect", "Error", MB_OK);
	}

	//��ȡ�������ָ�����
	hr = m_lpITS->GetFolder(_bstr_t("\\"), &m_lpRootFolder);
	if (FAILED(hr))
	{
		MessageBoxA(NULL, "GetFolder", "Error", MB_OK);
	}
}


CMyTaskSchedule::~CMyTaskSchedule()
{
	if (m_lpITS)
	{
		m_lpITS->Release();
	}
	if (m_lpRootFolder)
	{
		m_lpRootFolder->Release();
	}
	// ж��COM
	::CoUninitialize();
}


//************************************************************
// ��������: NewTask 
// ����˵��: �����ƻ�����
// ��	 ��: GuiShou
// ʱ	 ��: 2019/1/22
// ��	 ��: char* lpszTaskName    �ƻ�������
// ��	 ��: char* lpszProgramPath �ƻ�����·��
// ��	 ��: char* lpszParameters  �ƻ��������
// ��	 ��: char* szAuthor        �ƻ���������
// �� �� ֵ: BOOL�Ƿ�ɹ�
//************************************************************
BOOL CMyTaskSchedule::NewTask(char* lpszTaskName, char* lpszProgramPath, char* lpszParameters, char* lpszAuthor)
{
	if(NULL == m_lpRootFolder)
	{
		return FALSE;
	}
	// ���������ͬ�ļƻ�������ɾ��
	Delete(lpszTaskName);
	// �����������������������
	ITaskDefinition *pTaskDefinition = NULL;
	HRESULT hr = m_lpITS->NewTask(0, &pTaskDefinition);
	if(FAILED(hr))
	{
		MessageBoxA(0, "ITaskService::NewTask", "ERROR", 0);
		return FALSE;
	}

	/* ����ע����Ϣ */
	IRegistrationInfo *pRegInfo = NULL;
	CComVariant variantAuthor(NULL);
	variantAuthor = lpszAuthor;
	hr = pTaskDefinition->get_RegistrationInfo(&pRegInfo);
	if(FAILED(hr))
	{
		MessageBoxA(0, "pTaskDefinition::get_RegistrationInfo", "ERROR", 0);
		return FALSE;
	}
	// ����������Ϣ
	hr = pRegInfo->put_Author(variantAuthor.bstrVal);
	pRegInfo->Release();

	/* ���õ�¼���ͺ�����Ȩ�� */
	IPrincipal *pPrincipal = NULL;
	hr = pTaskDefinition->get_Principal(&pPrincipal);
	if(FAILED(hr))
	{
		MessageBoxA(0, "pTaskDefinition::get_Principal", "ERROR", 0);
		return FALSE;
	}
	// ���õ�¼����
	hr = pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
	// ��������Ȩ��
	// ���Ȩ��
	hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);  
	pPrincipal->Release();

	/* ����������Ϣ */
	ITaskSettings *pSettting = NULL;
	hr = pTaskDefinition->get_Settings(&pSettting);
	if(FAILED(hr))
	{
		MessageBoxA(0, "pTaskDefinition::get_Settings", "ERROR", 0);
		return FALSE;
	}
	// ����������Ϣ
	hr = pSettting->put_StopIfGoingOnBatteries(VARIANT_FALSE);
	hr = pSettting->put_DisallowStartIfOnBatteries(VARIANT_FALSE);
	hr = pSettting->put_AllowDemandStart(VARIANT_TRUE);
	hr = pSettting->put_StartWhenAvailable(VARIANT_FALSE);
	hr = pSettting->put_MultipleInstances(TASK_INSTANCES_PARALLEL);
	pSettting->Release();

	/* ����ִ�ж��� */
	IActionCollection *pActionCollect = NULL;
	hr = pTaskDefinition->get_Actions(&pActionCollect);
	if(FAILED(hr))
	{
		MessageBoxA(0, "pTaskDefinition::get_Actions", "ERROR", 0);
		return FALSE;
	}
	IAction *pAction = NULL;
	// ����ִ�в���
	hr = pActionCollect->Create(TASK_ACTION_EXEC, &pAction);
	pActionCollect->Release();

	/* ����ִ�г���·���Ͳ��� */
	CComVariant variantProgramPath(NULL);
	CComVariant variantParameters(NULL);
	IExecAction *pExecAction = NULL;
	hr = pAction->QueryInterface(IID_IExecAction, (PVOID *)(&pExecAction));
	if(FAILED(hr))
	{
		pAction->Release();
		MessageBoxA(0, "IAction::QueryInterface", "ERROR", 0);

		return FALSE;
	}
	pAction->Release();
	// ���ó���·���Ͳ���
	variantProgramPath = lpszProgramPath;
	variantParameters = lpszParameters;
	pExecAction->put_Path(variantProgramPath.bstrVal);
	pExecAction->put_Arguments(variantParameters.bstrVal);
	pExecAction->Release();

	/* ������������ʵ���û���½������ */
	ITriggerCollection *pTriggers = NULL;
	hr = pTaskDefinition->get_Triggers(&pTriggers);
	if (FAILED(hr))
	{
		MessageBoxA(0, "pTaskDefinition::get_Triggers", "ERROR", 0);
		return FALSE;
	}
	// ����������
	ITrigger *pTrigger = NULL;
	//�û���¼ʱ����
	hr = pTriggers->Create(TASK_TRIGGER_LOGON, &pTrigger);
	if (FAILED(hr))
	{
		MessageBoxA(0, "ITaskFolder::Create", "ERROR", 0);
		return FALSE;
	}

	/* ע������ƻ�  */
	IRegisteredTask *pRegisteredTask = NULL;
	CComVariant variantTaskName(NULL);
	variantTaskName = lpszTaskName;
	hr = m_lpRootFolder->RegisterTaskDefinition(variantTaskName.bstrVal,
		pTaskDefinition,
		TASK_CREATE_OR_UPDATE,
		_variant_t(),
		_variant_t(),
		TASK_LOGON_INTERACTIVE_TOKEN,
		_variant_t(""),
		&pRegisteredTask);
	if(FAILED(hr))
	{
		pTaskDefinition->Release();
		MessageBoxA(0, "ITaskFolder::RegisterTaskDefinition", "ERROR", 0);
		return FALSE;
	}
	pTaskDefinition->Release();
	pRegisteredTask->Release();

	return TRUE;

}


//************************************************************
// ��������: Delete 
// ����˵��: ɾ���ƻ�����
// ��	 ��: GuiShou
// ʱ	 ��: 2019/1/22
// ��	 ��: char * lpszTaskName �ƻ�������
// �� �� ֵ: BOOL�Ƿ�ɹ�
//************************************************************
BOOL CMyTaskSchedule::Delete(char * lpszTaskName)
{
	if (NULL==m_lpRootFolder)
	{
		return FALSE;
	}

	CComVariant variantTaskName(NULL);
	variantTaskName = lpszTaskName;
	HRESULT hr = m_lpRootFolder->DeleteFolder(variantTaskName.bstrVal, 0);
	if (FAILED(hr))
	{
		return FALSE;
	}
	return TRUE;
}



//************************************************************
// ��������: IsTaskValid 
// ����˵��: �ж�ָ���ƻ������Ƿ���
// ��	 ��: GuiShou
// ʱ	 ��: 2019/1/22
// ��	 ��: char * lpszTaskName �ƻ�������
// �� �� ֵ: BOOL�Ƿ�ɹ�
//************************************************************
BOOL CMyTaskSchedule::IsTaskValid(char * lpszTaskName)
{
	if (NULL == m_lpRootFolder)
	{
		return FALSE;
	}
	HRESULT hr = S_OK;
	CComVariant variantTaskName(NULL);
	CComVariant variantEnable(NULL);
	variantTaskName = lpszTaskName;                     // ����ƻ�����
	IRegisteredTask *pRegisteredTask = NULL;
	// ��ȡ����ƻ�
	hr = m_lpRootFolder->GetTask(variantTaskName.bstrVal, &pRegisteredTask);
	if (FAILED(hr) || (NULL == pRegisteredTask))
	{
		return FALSE;
	}
	pRegisteredTask->Release();

	return TRUE;
}


//************************************************************
// ��������: Run 
// ����˵��: ����ָ������ƻ�
// ��	 ��: GuiShou
// ʱ	 ��: 2019/1/22
// ��	 ��: char * lpszTaskName �ƻ�������  char * lpszParam����
// �� �� ֵ: BOOL�Ƿ�ɹ�
//************************************************************
BOOL CMyTaskSchedule::Run(char * lpszTaskName, char * lpszParam)
{
	if (NULL == m_lpRootFolder)
	{
		return FALSE;
	}
	HRESULT hr = S_OK;
	CComVariant variantTaskName(NULL);
	CComVariant variantParameters(NULL);
	variantTaskName = lpszTaskName;
	variantParameters = lpszParam;

	// ��ȡ����ƻ�
	IRegisteredTask *pRegisteredTask = NULL;
	hr = m_lpRootFolder->GetTask(variantTaskName.bstrVal, &pRegisteredTask);
	if (FAILED(hr) || (NULL == pRegisteredTask))
	{
		return FALSE;
	}
	// ����
	hr = pRegisteredTask->Run(variantParameters, NULL);
	if (FAILED(hr))
	{
		pRegisteredTask->Release();
		return FALSE;
	}
	pRegisteredTask->Release();

	return TRUE;
}


//************************************************************
// ��������: IsEnable 
// ����˵��: �ж�ָ������ƻ��Ƿ�����
// ��	 ��: GuiShou
// ʱ	 ��: 2019/1/22
// ��	 ��: char * lpszTaskName �ƻ�������
// �� �� ֵ: BOOL�Ƿ�ɹ�
//************************************************************
BOOL CMyTaskSchedule::IsEnable(char * lpszTaskName)
{
	if (NULL == m_lpRootFolder)
	{
		return FALSE;
	}
	HRESULT hr = S_OK;
	CComVariant variantTaskName(NULL);
	CComVariant variantEnable(NULL);
	variantTaskName = lpszTaskName;                     // ����ƻ�����
	IRegisteredTask *pRegisteredTask = NULL;
	// ��ȡ����ƻ�
	hr = m_lpRootFolder->GetTask(variantTaskName.bstrVal, &pRegisteredTask);
	if (FAILED(hr) || (NULL == pRegisteredTask))
	{
		return FALSE;
	}
	// ��ȡ�Ƿ��Ѿ�����
	pRegisteredTask->get_Enabled(&variantEnable.boolVal);
	pRegisteredTask->Release();
	if (ATL_VARIANT_FALSE == variantEnable.boolVal)
	{
		return FALSE;
	}

	return TRUE;
}


//************************************************************
// ��������: SetEnable 
// ����˵��: ����ָ������ƻ��Ƿ��������ǽ���
// ��	 ��: GuiShou
// ʱ	 ��: 2019/1/22
// ��	 ��: char * lpszTaskName �ƻ������� bEnable�Ƿ�����
// �� �� ֵ: BOOL�Ƿ�ɹ�
//************************************************************
BOOL CMyTaskSchedule::SetEnable(char * lpszTaskName, BOOL bEnable)
{
	if (NULL == m_lpRootFolder)
	{
		return FALSE;
	}
	HRESULT hr = S_OK;
	CComVariant variantTaskName(NULL);
	CComVariant variantEnable(NULL);
	variantTaskName = lpszTaskName;                     // ����ƻ�����
	variantEnable = bEnable;                            // �Ƿ�����
	IRegisteredTask *pRegisteredTask = NULL;
	// ��ȡ����ƻ�
	hr = m_lpRootFolder->GetTask(variantTaskName.bstrVal, &pRegisteredTask);
	if (FAILED(hr) || (NULL == pRegisteredTask))
	{
		return FALSE;
	}
	// �����Ƿ�����
	pRegisteredTask->put_Enabled(variantEnable.boolVal);
	pRegisteredTask->Release();

	return TRUE;
}


