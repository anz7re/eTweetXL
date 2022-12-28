Attribute VB_Name = "CtrlBox_INFO"
Public Function AppTag()

AppTag = "Control Box+"

End Function
Public Function AppLicense()

nl = vbNewLine

AppLicense = "License Information:" & _
nl & nl & "Copyright (C) 2022-present, Autokit Technology." & _
nl & nl & "Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:" & _
nl & nl & "1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer." & _
nl & nl & "2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution." & _
nl & nl & "3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission." & _
nl & nl & "THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS 'AS IS' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO," & _
nl & "THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES" & _
nl & "(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)" & _
nl & "HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE."

End Function
Public Function AppTitle()

AppTitle = Replace(CTRLBOX.Caption, CtrlBox_INFO.AppTag, vbNullString)
AppTitle = Replace(AppTitle, " - ", vbNullString)

End Function
