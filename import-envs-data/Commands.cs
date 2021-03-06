﻿using System;
using System.Linq;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Collections.Generic;

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

            // Cria um StreamReader para ler o arquivo 'ida.txt', que contém os
            // dados para inserir os blocos.
            var fileData = new StreamReader(curDwgPath + "\\ida.txt");

            // Diretório dos arquivos '.dwg' dos blocos. Pois, caso o bloco
            // não exista no desenho atual, o programa irá procurar os blocos
            // em arquivos '.dwg' externos.
            String blkPath = curDwgPath;
            
            // Armazenará uma linha que será lida do 'ida.txt'
            string[] sFileLine;

            // Informações dos blocos, que serão lidos do 'ida.txt'
            string sBlkId;
            string sBlkName;
            Point3d ptBlkOrigin;
            double dbBlkRot;
            string sLayer;
            string sColor;
            ObjectId idBlkTblRec = ObjectId.Null;
            BlockTableRecord blkTblRec = null;
            string[] sBlkAtts;
            string[] oneAtt;
            string blkTag;
            string blkValue;
            Dictionary<string, string> dicBlkAtts = new Dictionary<string, string>();

            using (var tr = db.TransactionManager.StartTransaction())
            {

                BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                
                // O ModelSpace será usado para gravar em seu Extension Dictionary os Id's e os handles dos blocos.
                DBObject dbModelSpace = tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                using (BlockTableRecord acBlkTblRec = (BlockTableRecord)tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead))
                {
                    int contId = 0;
                    while (!fileData.EndOfStream)
                    {
                        // sFileLine é uma array que lê uma linha do 'ida.txt'
                        // tendo como separador de colunas ';'
                        sFileLine = fileData.ReadLine().Split(';');

                        // Atribui os parâmetros de cada bloco que será inserido
                        // baseado no 'ida.txt'
                        sBlkId = Convert.ToString(contId); // O id não é declarado no 'ida.txt' pois este arquivo vem do matlab.
                        sBlkName = sFileLine[0];
                        dbBlkRot = (Math.PI / 180) * Convert.ToDouble(sFileLine[1]); // Converte graus para radiano
                        // Aqui é usado um Point3d pois é requisitado para criar um 'BlockReference' e não um Point2d.
                        ptBlkOrigin = new Point3d(Convert.ToDouble(sFileLine[2]), Convert.ToDouble(sFileLine[3]), 0);
                        sLayer = sFileLine[4];
                        sColor = sFileLine[5];
                        sBlkAtts = sFileLine[6].Split(new string[] { "//" }, StringSplitOptions.None);

                        foreach(var sBlkAtt in sBlkAtts)
                        {
                            oneAtt = sBlkAtt.Split(new string[] { "::" }, StringSplitOptions.None);
                            blkTag = oneAtt[0];
                            blkValue = oneAtt[1];

                            dicBlkAtts.Add(blkTag, blkValue);
                        }

                        // Se o bloco não existe no desenho atual
                        if (!acBlkTbl.Has(sBlkName))
                        {
                            try
                            {
                                using (var blkDb = new Database(false, true))
                                {
                                    // Lê o '.dwg' do bloco, baseado no diretório especificado.
                                    blkDb.ReadDwgFile(blkPath + "/" + sBlkName + ".dwg", FileOpenMode.OpenForReadAndAllShare, true, "");

                                    // E então insere o bloco no banco de dados do desenho atual.
                                    // Mas ainda não está inserido no desenho.
                                    idBlkTblRec = db.Insert(sBlkName, blkDb, true); // Este método retorna o id do bloco.
                                }
                            }
                            catch // Expressão para pegar erros.
                            {

                            }
                        }
                        else
                        {
                            idBlkTblRec = acBlkTbl[sBlkName];
                        }

                        if (idBlkTblRec != ObjectId.Null)
                        {
                            blkTblRec = idBlkTblRec.GetObject(OpenMode.ForWrite) as BlockTableRecord;

                            using (var trColor = db.TransactionManager.StartOpenCloseTransaction())
                            {
                                // Altera a cor para "ByBlock"
                                foreach (ObjectId oId in blkTblRec)
                                {
                                    Entity ent = (Entity)trColor.GetObject(oId, OpenMode.ForWrite);
                                    ent.ColorIndex = 0;
                                }
                                trColor.Commit();
                            }

                            // Aqui o bloco será adicionado ao desenho e serão gravados 
                            // seu id (identificação dada pelo programa atual, começando em 0 e incrementado 1 a 1)
                            // pareado ao seu handle, para o uso dos outros comandos.
                            using (BlockReference acBlkRef = new BlockReference(ptBlkOrigin, idBlkTblRec))
                            {
                                BlockTableRecord acCurSpaceBlkTblRec;
                                acCurSpaceBlkTblRec = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                                tr.AddNewlyCreatedDBObject(acBlkRef, true);
                                acBlkRef.Rotation = dbBlkRot;
                                acBlkRef.Layer = sLayer;
                                acBlkRef.ColorIndex = Convert.ToInt32(sColor);

                                // Início: Setar atributos do bloco ***********************************************************
                                if (blkTblRec.HasAttributeDefinitions)
                                {
                                    RXClass rxClass = RXClass.GetClass(typeof(AttributeDefinition));
                                    foreach (ObjectId idAttDef in blkTblRec)
                                    {
                                        if(idAttDef.ObjectClass == rxClass)
                                        {
                                            DBObject obj = tr.GetObject(idAttDef, OpenMode.ForRead);

                                            AttributeDefinition ad = obj as AttributeDefinition;
                                            AttributeReference ar = new AttributeReference();
                                            ar.SetAttributeFromBlock(ad, acBlkRef.BlockTransform);

                                            if (dicBlkAtts.ContainsKey(ar.Tag))
                                            {
                                                ar.TextString = dicBlkAtts[ar.Tag];
                                            }
                                            

                                            acBlkRef.AttributeCollection.AppendAttribute(ar);
                                            tr.AddNewlyCreatedDBObject(ar, true);
                                        }
                                    }

                                    // Setar propriedades dos blocos dinâmicos
                                    if (acBlkRef.IsDynamicBlock)
                                    {
                                        DynamicBlockReferencePropertyCollection pc = acBlkRef.DynamicBlockReferencePropertyCollection;
                                        foreach(DynamicBlockReferenceProperty prop in pc)
                                        {
                                            if(dicBlkAtts.ContainsKey(prop.PropertyName))
                                            {
                                                // Propriedade de distância
                                                if(prop.PropertyTypeCode == 1)
                                                {
                                                    prop.Value = Convert.ToDouble(dicBlkAtts[prop.PropertyName]);
                                                }
                                                // Propriedade visibilidade
                                                else if(prop.PropertyTypeCode == 5)
                                                {
                                                    prop.Value = dicBlkAtts[prop.PropertyName];
                                                }
                                                
                                            }
                                        }
                                    }
                                }
                                // Fim: Setar atributos do bloco ************************************************************

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
                        dicBlkAtts.Clear();
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
            var fileData = new StreamReader(curDwgPath + "\\let.txt");
            // ***************************************************************************************************************************

            string[] sLine;
            List<string> lLine = new List<string>();
            
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
                int lineColor = -1;

                while (!fileData.EndOfStream)
                {
                    sLine = fileData.ReadLine().Split(';');
                    indexVert = 0;

                    foreach(string value in sLine)
                    {
                        if(value != "")
                        {
                            lLine.Add(value);
                        }
                    }

                    lineColor = Convert.ToInt32(lLine.Last());
                    lLine.RemoveAt(lLine.Count - 1);

                    using (Polyline oPLine = new Polyline())
                    {
                        // address pode ser um index ou um ponto
                        foreach(string address in lLine)
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

                                oPLine.AddVertexAt(indexVert, ptVert, 0, 0, 0);
                                indexVert++;
                            }
                        }
                        oPLine.ColorIndex = lineColor;
                        BlkTblRec.AppendEntity(oPLine);
                        tr.AddNewlyCreatedDBObject(oPLine, true);
                    }
                    lLine.Clear();
                }
                tr.Commit();
            }
        }


        // Comando que lê dados dos blocos selecionados pelo usuário
        // e exporta um arquivo '.txt' com esses dados.
        [CommandMethod("LEBLOCK")]
        public void ReadBlocksData()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Diretório do desenho atual que chamou o comando.
            string curDwgPath = AcAp.GetSystemVariable("DWGPREFIX").ToString();

            // Arquivo onde será escrito os dados obtidos.
            StreamWriter fileOut = new StreamWriter(curDwgPath + "\\blocksData.txt");

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
                        // Início: Pega o id guardado em seu XDic******************************************************
                        var blkRef = (BlockReference)tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead);

                        // Se o bloco não tem XDic, cria-o e já grava o dado
                        if(blkRef.ExtensionDictionary == ObjectId.Null)
                        {
                            blkRef.UpgradeOpen();
                            RecOnXDict(blkRef, "id", DxfCode.XTextString, dbExt.Count.ToString(), tr);
                            RecOnXDict(dbModelSpace, dbExt.Count.ToString(), DxfCode.Handle, blkRef.Handle, tr);
                        }

                        var blkDic = (DBDictionary)tr.GetObject(blkRef.ExtensionDictionary, OpenMode.ForRead);
                        var xRecBlk = (Xrecord)tr.GetObject(blkDic.GetAt("id"), OpenMode.ForRead);
                        ResultBuffer rb = xRecBlk.Data;
                        TypedValue[] xRecData = rb.AsArray();
                        string sBlkId = xRecData[0].Value.ToString();
                        // Fim: Pega o id guardado em seu XDic*************************************************************************************

                        // Início: checa se o id do bloco confere com o handle registrado no MS ***************************************************
                        var blkRef2 = GetRefBlkFromIndex(db, dbExt, sBlkId);

                        if(blkRef.Id != blkRef2.Id)
                        {
                            RecOnXDict(blkRef, "id", DxfCode.XTextString, dbExt.Count.ToString(), tr);
                            RecOnXDict(dbModelSpace, dbExt.Count.ToString(), DxfCode.Handle, blkRef.Handle, tr);
                        }
                        // Fim: checa se o id do bloco confere com o handle registrado no MS ******************************************************

                        string blkName = blkRef.Name;

                        // Se é bloco dinamico, a forma de pegar o nome do bloco
                        // é desse jeito:
                        BlockTableRecord realBlock = null;
                        string atts = "";
                        if (blkRef.IsDynamicBlock)
                        {
                            realBlock = (BlockTableRecord)tr.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForRead);
                            blkName = realBlock.Name;

                            // Le os atributos do bloco
                            if (realBlock.HasAttributeDefinitions)
                            {
                                var blkRefAtts = (AttributeCollection)blkRef.AttributeCollection;
                                RXClass rxClass = RXClass.GetClass(typeof(AttributeReference));
                                foreach (ObjectId idAttDef in blkRefAtts)
                                {
                                    if (idAttDef.ObjectClass == rxClass)
                                    {
                                        AttributeReference obj = (AttributeReference)tr.GetObject(idAttDef, OpenMode.ForRead);

                                        atts += obj.Tag + "::" + obj.TextString + "//";
                                    }
                                }
                            }
                        }

                        double blkRot = blkRef.Rotation * (180/Math.PI);
                        string blkX = blkRef.Position.X.ToString("n2");
                        string blkY = blkRef.Position.Y.ToString("n2");
                        string blkLayer = blkRef.Layer;
                        string blkColor = Convert.ToString(blkRef.ColorIndex);

                        fileOut.WriteLine(sBlkId + ";" + blkName + ";" + blkRot + ";" + blkX + ";" + blkY + ";" 
                            + blkLayer + ";" + blkColor + ";" + atts + ";");
                    }
                }

                fileOut.Close();
                tr.Commit();
            }
        }

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

        [CommandMethod("MoveBlock")]
        public void MoveBlock()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            // LOcal do arquivo de dados: O diretório do .dwg onde está send
            // executado o programa.
            string curDwgPath = AcAp.GetSystemVariable("DWGPREFIX").ToString();

            // Cria um StreamReader para ler o arquivo 'moveblocks.txt', que contém 
            // os dados para mover os blocos.
            var fileData = new StreamReader(curDwgPath + "\\moveBlocks.txt");

            // Armazenará uma linha que será lida do 'moveBlocks.txt'
            string[] sFileLine;

            string blockId;
            string[] coords;
            string coordX;
            string coordY;
            Point3d oldPos;
            Point3d newPos;
            Vector3d vector;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                while (!fileData.EndOfStream)
                {
                    sFileLine = fileData.ReadLine().Split(';');
                    blockId = sFileLine[0];
                    coords = sFileLine[1].Split(',');
                    coordX = coords[0];
                    coordY = coords[1];

                    DBObject dbModelSpace = tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                    ObjectId extId = dbModelSpace.ExtensionDictionary;
                    DBDictionary dbExt = (DBDictionary)tr.GetObject(extId, OpenMode.ForRead);

                    BlockReference blkRef = GetRefBlkFromIndex(db, dbExt, blockId);
                    BlockReference blkRefWrite = (BlockReference)tr.GetObject(blkRef.Id, OpenMode.ForWrite);

                    oldPos = blkRef.Position;
                    newPos = new Point3d(Convert.ToDouble(coordX), Convert.ToDouble(coordY), 0);
                    vector = oldPos.GetVectorTo(newPos);

                    blkRefWrite.TransformBy(Matrix3d.Displacement(vector));
                }
                tr.Commit();
            }
        }
    }
}
