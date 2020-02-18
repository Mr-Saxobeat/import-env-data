using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AcadPlugin
{
    public class Commands
    {
        [CommandMethod("IDA")]
        public void ImportEnvData()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Local do arquivo de dados: O diretório do .dwg onde está sendo 
            // executado o programa.
            string curDwgPath = AcAp.GetSystemVariable("DWGPREFIX").ToString();

            // Cria um StreamReader para ler o arquivo 'ida.csv', que contém os
            // dados para inserir os blocos.
            var fileData = new StreamReader(curDwgPath + "\\ida.csv");

            // Diretório dos arquivos '.dwg' dos blocos. Pois, caso o bloco
            // não exista no desenho atual, o programa irá procurar os blocos
            // em arquivos '.dwg' externos.
            String blkPath = curDwgPath;
            
            // Armazenará uma linha que será lida do 'ida.csv'
            string[] sFileLine;

            // Informações dos blocos, que serão lidos do 'ida.csv'
            string sBlkId;
            string sBlkName;
            //string[] sBlkCoord;
            Point3d ptBlkOrigin;
            string sBlkRot;
            ObjectId blkId = ObjectId.Null;

            using (var tr = db.TransactionManager.StartTransaction())
            {

                BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                // O ModelSpace será usado para gravar em seu Extension Dictionary os Id's e os handles dos blocos.
                DBObject dbModelSpace = tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                using (var acBlkTblRec = new BlockTableRecord())
                {
                    int contId = 0;
                    while (!fileData.EndOfStream)
                    {
                        // sFileLine é uma array que lê uma linha do 'ida.csv'
                        // tendo como separador de colunas ';'
                        sFileLine = fileData.ReadLine().Split(';');

                        // Atribui os parâmetros de cada bloco que será inserido
                        // baseado no 'ida.csv'
                        sBlkId = Convert.ToString(contId); // O id não é declarado no 'ida.csv' pois este arquivo vem do matlab.
                        sBlkName = sFileLine[0];
                        sBlkRot = sFileLine[1];
                        // Aqui é usado um Point3d pois é requisitado para criar um 'BlockReference' e não um Point2d.
                        ptBlkOrigin = new Point3d(Convert.ToDouble(sFileLine[2]), Convert.ToDouble(sFileLine[3]), 0);




                        // ******************************************************************************************************************************************
                        // ******************************************************************************************************************************************
                        // Falta pegar os valores dos atributos dos blocos (que é o último argumento dado no 'ida.csv').
                        // ******************************************************************************************************************************************
                        // ******************************************************************************************************************************************



                        // Se o bloco não existe no desenho atual
                        if (!acBlkTbl.Has(sBlkName))
                        {
                            try
                            {
                                using (var blkDb = new Database(false, true))
                                {
                                    // Lêo '.dwg' do bloco, baseado no diretório especificado.
                                    blkDb.ReadDwgFile(blkPath + "/" + sBlkName + ".dwg", FileOpenMode.OpenForReadAndAllShare, true, "");

                                    // E então insere o bloco no banco de dados do desenho atual.
                                    // Mas ainda não está inserido no desenho.
                                    blkId = db.Insert(sBlkName, blkDb, true); // Este método retorna o id do bloco.
                                }
                            }
                            catch // Expressão para pegar erros.
                            {
                                
                            }
                        }

                        if (blkId != ObjectId.Null)
                        {
                            // Aqui o bloco será adicionado ao desenho e serão gravados 
                            // seu id (identificação dada pelo programa atual, começando em 0 e incrementado 1 a 1)
                            // pareado ao seu handle, para o uso dos outros comandos.
                            using (var acBlkRef = new BlockReference(ptBlkOrigin, blkId))
                            {
                                BlockTableRecord acCurSpaceBlkTblRec;
                                acCurSpaceBlkTblRec = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                                tr.AddNewlyCreatedDBObject(acBlkRef, true);

                                Entity eBlk = (Entity)tr.GetObject(acBlkRef.Id, OpenMode.ForRead);
                                DBObject blkDb = (DBObject)tr.GetObject(acBlkRef.Id, OpenMode.ForRead);

                                // Grava o id(dado pelo programa - base 0) e o handle do bloco
                                // no Extension Dictionary do ModelSpace.
                                RecOnXDict(dbModelSpace, sBlkId, DxfCode.Handle, eBlk.Handle, tr);

                                // Grava o id do bloco em seu próprio Extension Dictionary, para fazer
                                // uma 'busca reversa' no XDic do ModelSpace depois.
                                RecOnXDict(blkDb, "id", DxfCode.XTextString, sBlkId, tr);
                            }
                        }
                        contId++;
                    }
                }
                fileData.Close();
                tr.Commit();
            }
        }

        // Função para gravar os dados no Extension Dictionary do objeto dado.
        public void RecOnXDict(DBObject dbObj, string location, DxfCode dxfC, dynamic data, Transaction tr)
        {
            // Pega o Id do XDic do objeto.
            ObjectId extId = dbObj.ExtensionDictionary;

            // Se não existe um XDic, cria-o
            if (extId == ObjectId.Null)
            {
                dbObj.CreateExtensionDictionary();
                extId = dbObj.ExtensionDictionary;
            }

            // Pega o XDic a partir do seu Id.
            DBDictionary dbExt = (DBDictionary)tr.GetObject(extId, OpenMode.ForWrite);

            // Cria um XRecord, que guardará a informação
            Xrecord xRec = new Xrecord();
            ResultBuffer rb = new ResultBuffer();
            rb.Add(new TypedValue((int)dxfC, data));

            xRec.Data = rb;

            // Adiciona a informação no XDic do objeto.
            dbExt.SetAt(location, xRec);
            tr.AddNewlyCreatedDBObject(xRec, true);
        }

        // Comando que liga os eletrodutos a partir do arquivo
        // 'let.csv' que se encontra no diretório do desenho atual.
        // O arquivo 'let.csv' é exportado pelo matlab com o formato:
        // id1;id2;id3; OU id1;pontoX,pontoY;id2;
        // e qualquer combinação de ids com pontos
        // (note que as coordenadas de um ponto são separados por vírgula).
        [CommandMethod("ELET")]
        public void ConnectConduits()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Local do arquivo de dados 
            // ***************************************************************************************************************************
            //string curDwgPath = Directory.GetCurrentDirectory();
            string curDwgPath = AcAp.GetSystemVariable("DWGPREFIX").ToString();
            var fileData = new StreamReader(curDwgPath + "\\let.csv");
            // ***************************************************************************************************************************

            string[] sLine;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                BlockTable BlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                BlockTableRecord BlkTblRec = tr.GetObject(BlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                DBObject dbModelSpace = tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                ObjectId extId = dbModelSpace.ExtensionDictionary;
                DBDictionary dbExt = (DBDictionary)tr.GetObject(extId, OpenMode.ForRead);
                Point3d ptAux;
                Point2d ptVert;
                int indexVert;

                while (!fileData.EndOfStream)
                {
                    sLine = fileData.ReadLine().Split(';');
                    indexVert = 0;

                    using(Polyline oPLine = new Polyline())
                    {
                        // address pode ser um index ou um ponto
                        foreach(string address in sLine)
                        {
                            if(address != "")
                            {
                                // Se tem vírgula é um ponto, onde suas coordenadas
                                // são separadas por uma vírgula.
                                if (address.Contains(','))
                                {
                                    string[] coords = address.Split(',');
                                    ptVert = new Point2d(Convert.ToDouble(coords[0]), Convert.ToDouble(coords[1]));
                                }
                                // Se não tem vírgula, é um index, 
                                // na qual é usado o método 'GetRefBlkFromIndex' 
                                // para pegar o ponto do bloco.
                                else
                                {
                                    ptAux = GetRefBlkFromIndex(db, dbExt, address).Position;
                                    ptVert = new Point2d(ptAux.X, ptAux.Y);
                                }

                                //oPLine.SetPointAt(indexPoint, ptVert);
                                oPLine.AddVertexAt(indexVert, ptVert, 0, 0, 0);
                                indexVert++;
                            }
                        }
                        BlkTblRec.AppendEntity(oPLine);
                        tr.AddNewlyCreatedDBObject(oPLine, true);
                    }
                }
                tr.Commit();
            }
        }


        // Comando que lê dados dos blocos selecionados pelo usuário
        // e exporta um arquivo '.csv' com esses dados.
        [CommandMethod("LEBLOCK")]
        public void ReadBlocksData()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Diretório do desenho atual que chamou o comando.
            string curDwgPath = AcAp.GetSystemVariable("DWGPREFIX").ToString();

            // Arquivo onde será escrito os dados obtidos.
            StreamWriter fileOut = new StreamWriter(curDwgPath + "\\blocksData.csv");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var BlkTbl = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                var BlkTblRec = (BlockTableRecord)tr.GetObject(BlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var dbModelSpace = (DBObject)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                ObjectId extId = dbModelSpace.ExtensionDictionary;
                var dbExt = (DBDictionary)tr.GetObject(extId, OpenMode.ForRead);

                PromptSelectionResult acSSRes = ed.GetSelection();
                if (acSSRes.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSRes.Value;

                    foreach (SelectedObject selectedObject in acSSet)
                    {
                        // Pega o id guardado em seu XDic******************************************************
                        var blkRef = (BlockReference)tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead);

                        // Se o bloco não tem XDic, cria-o e já grava o dado tanto em seu próprio
                        // XDic como no XDic do ModelSpace.
                        if(blkRef.ExtensionDictionary == ObjectId.Null)
                        {
                            blkRef.UpgradeOpen();
                            RecOnXDict((DBObject)blkRef, "id", DxfCode.XTextString, dbExt.Count.ToString(), tr);
                            RecOnXDict(dbModelSpace, dbExt.Count.ToString(), DxfCode.Handle, blkRef.Handle, tr);
                        }

                        var blkDic = (DBDictionary)tr.GetObject(blkRef.ExtensionDictionary, OpenMode.ForRead);
                        var xRecBlk = (Xrecord)tr.GetObject(blkDic.GetAt("id"), OpenMode.ForRead);
                        ResultBuffer rb = xRecBlk.Data;
                        TypedValue[] xRecData = rb.AsArray();
                        string sBlkId = xRecData[0].Value.ToString();
                        //*************************************************************************************

                        string blkName = blkRef.Name;
                        double blkRot = blkRef.Rotation;
                        string blkX = blkRef.Position.X.ToString("n2");
                        string blkY = blkRef.Position.Y.ToString("n2");

                        //************************************************************************************************************************
                        // Falta Pegar o valor do atributo (que ainda nem foi setado) ************************************************************
                        //************************************************************************************************************************

                        fileOut.WriteLine(sBlkId + ";" + blkName + ";" + blkRot + ";" + blkX + ";" + blkY + ";");
                    }
                }
                fileOut.Close();
                tr.Commit();
            }
        }

        // Função para pegar o ReferenceBlock gravado no XDic do ModelSpace
        // a partir do id dado pelo programa.
        public BlockReference GetRefBlkFromIndex(Database db, DBDictionary dbExt, string idHn)
        {
            BlockReference blk;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Pega o handle a partir do id dado 
                ObjectId idId = dbExt.GetAt(idHn);
                Xrecord xRec = (Xrecord)tr.GetObject(idId, OpenMode.ForRead);
                ResultBuffer rb = xRec.Data;
                TypedValue[] tp = rb.AsArray();
                string hand = tp[0].Value as string;

                long ln = Convert.ToInt64(hand, 16);
                Handle hn = new Handle(ln);
                ObjectId id = db.GetObjectId(false, hn, 0);

                blk = (BlockReference)tr.GetObject(id, OpenMode.ForRead);

            }

            return blk;
        }


    }
}
