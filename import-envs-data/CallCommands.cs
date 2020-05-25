using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Collections.Generic;

namespace AcadPlugin
{
    public class CallCommands
    {
        [CommandMethod("AAA")]
        public void Call()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            string line;

            // Local do arquivo de dados: O diretório do .dwg onde está sendo 
            // executado o programa.
            string curDwgPath = AcAp.GetSystemVariable("DWGPREFIX").ToString();

            var commandsFile = new StreamReader(curDwgPath + "\\commands.txt");

            while (!commandsFile.EndOfStream)
            {
                line = commandsFile.ReadLine();
                if(line != "")
                {
                    try
                    {
                        ed.Command(line);
                    }
                    catch (System.Exception)
                    {
                        throw;
                    }
                }
            }

            commandsFile.Close();
        }
    }
}