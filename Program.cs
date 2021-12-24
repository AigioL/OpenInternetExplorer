using System;

var scriptType = Type.GetTypeFromCLSID(Guid.Parse("0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC"));
dynamic obj = Activator.CreateInstance(scriptType, false);
obj.Language = "vbscript";
const string content = "CreateObject(\"InternetExplorer.Application\").Visible=true";
obj.ExecuteStatement(content);