'************************************************************************
'Folder creations
Environment.Value("Images") = fn_FolderOperations("Images")
Environment.Value("Screenshots") = fn_FolderOperations("Screenshots")
Environment.Value("FunctionLog") = fn_FolderOperations("FunctionLog")
Environment.Value("EngineLog") = fn_FolderOperations("EngineLog")
'*************************************************************************

'********************************************************************************************
'class creating and using
'Set ClassLog = triggerClass
'ClassLog.fn_CreateLogFile "FunctionLog","FunctionLogPath"
'ClassLog.fn_CreateLogFile "EngineLog","EngineLogPath"
'ClassLog.fn_PrintnUpdateLogFile  "FunctionLogPath","Step1 of function log is successfull"
'ClassLog.fn_PrintnUpdateLogFile  "EngineLogPath","Step1 of engine log is successfull"
'********************************************************************************************
Call fn_CreateLogFile("FunctionLog","FunctionLogPath")
Call fn_CreateLogFile("EngineLog","EngineLogPath")
Call fn_PrintnUpdateLogFile("FunctionLogPath","Step1 of function log is successfull")