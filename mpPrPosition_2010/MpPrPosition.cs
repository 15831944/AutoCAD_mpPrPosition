#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Runtime;
using mpMsg;
using mpProductInt;
using ModPlus;
using Exception = System.Exception;

namespace mpPrPosition
{
    public class MpPrPosition : IExtensionApplication
    {
        [CommandMethod("ModPlus", "mpPrPosition", CommandFlags.UsePickSet)]
        public void AddPositions()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            //var filList = new[] { new TypedValue((int)DxfCode.Start, "INSERT") };
            //var filter = new SelectionFilter(filList);
            var opts = new PromptSelectionOptions();
            opts.Keywords.Add("Удалить");
            var kws = opts.Keywords.GetDisplayString(true);
            opts.MessageForAdding = "\nВыберите объекты, относящиеся к изделиям ModPlus: " + kws;
            opts.KeywordInput += delegate (object sender, SelectionTextInputEventArgs e)
            {
                if (e.Input.Equals("Удалить"))
                    DeletePositions();
            };
            //var res = ed.GetSelection(opts, filter);
            var res = ed.GetSelection(opts);
            if (res.Status != PromptStatus.OK)
                return;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Создаем список "элемент + марка" и список позиций
                    var elements = new List<string>();
                    var elementsExsist = new List<string>();
                    var positions = new List<string>();
                    var markTypes = new List<int>();
                    // Тип маркировки

                    foreach (var objectId in res.Value.GetObjectIds())
                    {
                        var ent = (Entity)tr.GetObject(objectId, OpenMode.ForWrite);
                        // Расширенные данные в виде спец.класса
                        var productInsert = (MpProductToSave)XDataHelpersForProducts.NewFromEntity(ent);
                        if (productInsert == null) continue;
                        // Теперь получаем сам продукт
                        var product = MpProduct.GetProductFromSaved(productInsert);
                        // Если есть данные и нет позиции
                        if (product != null)
                        {
                            int markType;
                            if (string.IsNullOrEmpty(product.Position.Trim()))
                            {
                                var element = product.GetNameByRule();
                                if (!element.Contains(product.BaseDocument.ShortName))
                                    element = product.BaseDocument.ShortName + " " + element;
                                // Если еще небыло
                                if (!elements.Contains(element))
                                {
                                    // Подсвечиваем
                                    ent.Highlight();
                                    // Запрос пользователю
                                    var pso = new PromptStringOptions("Элемент: " + element + ". Укажите позицию: ")
                                    {
                                        AllowSpaces = true
                                    };
                                    var pres = ed.GetString(pso);
                                    if (pres.Status != PromptStatus.OK)
                                    {
                                        ent.Unhighlight();
                                        continue;
                                    }
                                    elements.Add(element);
                                    positions.Add(pres.StringResult);
                                    product.Position = pres.StringResult;
                                    XDataHelpersForProducts.SaveDataToEntity(product.SetProductToSave(), ent, tr);
                                    // Маркировка
                                    AddPositionMarker(ent, 0, pres.StringResult, element, out markType);
                                    // Убираем подсветку
                                    ent.Unhighlight();
                                    markTypes.Add(markType);
                                }
                                else // Если уже был
                                {
                                    product.Position = positions[elements.IndexOf(element)];
                                    XDataHelpersForProducts.SaveDataToEntity(product.SetProductToSave(), ent, tr);
                                    // Маркировка
                                    AddPositionMarker(ent, markTypes[elements.IndexOf(element)], positions[elements.IndexOf(element)], element, out markType);
                                }
                            }
                            // Если позиция уже есть
                            else
                            {
                                // Подсвечиваем
                                ent.Highlight();
                                var element = product.GetNameByRule();
                                // Если еще небыло
                                if (!elementsExsist.Contains(element))
                                {
                                    // Маркировка
                                    AddPositionMarker(ent, 0, product.Position, element, out markType);
                                    elements.Add(element);
                                    elementsExsist.Add(element);
                                    markTypes.Add(markType);
                                }
                                else
                                {
                                    if(markTypes.Count > 0 & elements.Count > 0)
                                        AddPositionMarker(ent, markTypes[elements.IndexOf(element)], product.Position, element, out markType);
                                    else AddPositionMarker(ent, 0, product.Position, element, out markType);
                                }
                                // Убираем подсветку
                                ent.Unhighlight();
                            }
                        }
                    }
                }

                catch (Exception ex)
                {
                    MpExWin.Show(ex);
                }
                finally
                {
                    // В любом случае убираем подсветку со всех объектов
                    foreach (var objectId in res.Value.GetObjectIds())
                    {
                        var ent = (Entity)tr.GetObject(objectId, OpenMode.ForWrite);
                        ent.Unhighlight();
                    }
                    tr.Commit();
                }
            }
        }
        // Удаление позиций
        private static void DeletePositions()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = HostApplicationServices.WorkingDatabase;
            var opts = new PromptSelectionOptions
            {
                MessageForAdding = "\nВыберите объекты, относящиеся к изделиям ModPlus, для удаления позиции: "
            };
            var res = ed.GetSelection(opts);
            if (res.Status != PromptStatus.OK)
                return;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    foreach (var objectId in res.Value.GetObjectIds())
                    {
                        var ent = (Entity)tr.GetObject(objectId, OpenMode.ForWrite);
                        // Расширенные данные в виде спец.класса
                        var productInsert = (MpProductToSave)XDataHelpersForProducts.NewFromEntity(ent);
                        if (productInsert == null) continue;
                        // Теперь получаем сам продукт
                        var product = MpProduct.GetProductFromSaved(productInsert);
                        // Если есть данные и нет позиции
                        if (product != null)
                        {
                            product.Position = string.Empty;
                            XDataHelpersForProducts.SaveDataToEntity(product.SetProductToSave(), ent, tr);
                        }
                    }
                }

                catch (Exception ex)
                {
                    MpExWin.Show(ex);
                }
                finally
                {
                    // В любом случае убираем подсветку со всех объектов
                    foreach (var objectId in res.Value.GetObjectIds())
                    {
                        var ent = (Entity)tr.GetObject(objectId, OpenMode.ForWrite);
                        ent.Unhighlight();
                    }
                    tr.Commit();
                }
            }
        }

        /// <summary>
        /// Простановка маркировки позиции
        /// </summary>
        /// <param name="ent">Примитив</param>
        /// <param name="posTxt">Содержимое маркировки</param>
        /// <param name="type">Тип маркировки. 0 - нужно ставить новую марку, 1 - текст, 2 - выноска, 3 - ничего</param>
        /// <param name="element">Имя изделия для отображения в запросе</param>
        /// <param name="markType"> </param>
        private static void AddPositionMarker(Entity ent, int type, string posTxt, string element, out int markType)
        {
            markType = 3;
            if (type != 0)
            {
                if (type == 1)
                {
                    markType = 1;
                    AddTextMark(ent, posTxt);
                }
                else if (type == 2)
                {
                    markType = 2;
                    AddLeaderMark(ent, posTxt);
                }
                else if (type == 3)
                {
                    markType = 3;
                }
            }
            else
            {
                var pko = new PromptKeywordOptions("Выберите тип маркировки для изделия: " + element + " Позиция: " + posTxt + ": [Ничего/Текст/Выноска]", "Nothing Text Leader")
                {
                    AllowArbitraryInput = true,
                    AllowNone = false
                };
                var pr = AcApp.DocumentManager.MdiActiveDocument.Editor.GetKeywords(pko);
                if (pr.Status != PromptStatus.OK || string.IsNullOrEmpty(pr.StringResult))
                {
                    markType = 3; return;
                }
                // Далее в зависимости от выбранного условия
                switch (pr.StringResult)
                {
                    case "Nothing":
                        markType = 3;
                        break;
                    case "Text":
                        markType = 1;
                        AddTextMark(ent, posTxt);
                        break;
                    case "Leader":
                        markType = 2;
                        AddLeaderMark(ent, posTxt);
                        break;
                }
            }
        }
        private static void AddTextMark(Entity ent, string posTxt)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);
                // Получаем координаты объекта
                var ext = ent.GeometricExtents;
                // Получаем координату нужной точки (center)
                var pt = new Point3d(
                    (ext.MinPoint.X + ext.MaxPoint.X) / 2,
                    (ext.MinPoint.Y + ext.MaxPoint.Y) / 2,
                    (ext.MinPoint.Z + ext.MaxPoint.Z) / 2);
                // Создаем текст
                var txt = new DBText
                {
                    TextString = posTxt,
                    TextStyleId = db.Textstyle,
                    Justify = AttachmentPoint.MiddleCenter,
                    Position = pt,
                    AlignmentPoint = pt
                };

                txt.SetFromStyle();
                btr.AppendEntity(txt);
                tr.AddNewlyCreatedDBObject(txt, true);
                tr.TransactionManager.QueueForGraphicsFlush();
                tr.Commit();
            }
        }
        private static void AddLeaderMark(Entity ent, string posTxt)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            // Получаем координаты объекта
            var ext = ent.GeometricExtents;
            // Получаем координату нужной точки
            var pt = new Point3d(
                (ext.MinPoint.X + ext.MaxPoint.X) / 2,
                (ext.MinPoint.Y + ext.MaxPoint.Y) / 2,
                (ext.MinPoint.Z + ext.MaxPoint.Z) / 2);
            // Создаем текст
            var jig = new MLeaderJig
            {
                FirstPoint = pt,
                MlText = posTxt
            };
            var res = ed.Drag(jig);
            if (res.Status == PromptStatus.OK)
            {
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    btr.AppendEntity(jig.MLeader());
                    tr.AddNewlyCreatedDBObject(jig.MLeader(), true);
                    tr.Commit();
                }
            }
            doc.TransactionManager.QueueForGraphicsFlush();
        }

        public void Initialize()
        {
            ObjectContextMenu.Attach();
        }

        public void Terminate()
        {
            // nothing
        }
    }
    public class ObjectContextMenu
    {
        public static ContextMenuExtension MpPrPositionMarkCme;
        public static void Attach()
        {
            if (MpPrPositionMarkCme == null)
            {
                MpPrPositionMarkCme = new ContextMenuExtension();
                var miEnt = new MenuItem("MP:Выноска с маркой");
                miEnt.Click += SendCommand;
                MpPrPositionMarkCme.MenuItems.Add(miEnt);
                //ADD the popup item
                MpPrPositionMarkCme.Popup += contextMenu_Popup;

                var rxcEnt = RXObject.GetClass(typeof(Entity));
                Autodesk.AutoCAD.ApplicationServices.Application.AddObjectContextMenuExtension(rxcEnt, MpPrPositionMarkCme);
            }
        }
        public static void Detach()
        {
            if (MpPrPositionMarkCme != null)
            {
                var rxcEnt = RXObject.GetClass(typeof(Entity));
                Autodesk.AutoCAD.ApplicationServices.Application.RemoveObjectContextMenuExtension(rxcEnt, MpPrPositionMarkCme);
            }
        }

        public static void SendCommand(object o, EventArgs e)
        {
            AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_.MPLEADERWITHMARKFORPRODUCT ", true, false, false);
        }
        [CommandMethod("ModPlus", "mpLeaderWithMarkForProduct", CommandFlags.UsePickSet)]
        public static void AddLeaderWithMark()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            var res = ed.SelectImplied();
            if (res.Status != PromptStatus.OK)
                return;
            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        foreach (var objectId in res.Value.GetObjectIds())
                        {
                            var ent = (Entity)tr.GetObject(objectId, OpenMode.ForWrite);
                            // Расширенные данные в виде спец.класса
                            var productInsert = (MpProductToSave)XDataHelpersForProducts.NewFromEntity(ent);
                            if (productInsert == null) continue;
                            // Теперь получаем сам продукт
                            var product = MpProduct.GetProductFromSaved(productInsert);
                            // Если есть данные и нет позиции
                            if (product != null)
                            {
                                var mark = product.GetNameByRule();
                                // Если в имени нет ShortName, то добавим
                                if (!mark.Contains(product.BaseDocument.ShortName))
                                    mark = product.BaseDocument.ShortName + " " + mark;
                                // Если есть длина
                                if (product.Length != null)
                                    mark += " L=" + product.Length;
                                AddLeaderMark(ent, mark);
                            }
                        }
                    }

                    catch (Exception ex)
                    {
                        MpExWin.Show(ex);
                    }
                    finally
                    {
                        tr.Commit();
                    }
                }
            }
        }
        private static void AddLeaderMark(Entity ent, string posTxt)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            // Получаем координаты объекта
            var ext = ent.GeometricExtents;
            // Получаем координату нужной точки
            var pt = new Point3d(
                (ext.MinPoint.X + ext.MaxPoint.X) / 2,
                (ext.MinPoint.Y + ext.MaxPoint.Y) / 2,
                (ext.MinPoint.Z + ext.MaxPoint.Z) / 2);
            // Создаем текст
            var jig = new MLeaderJig
            {
                FirstPoint = pt,
                MlText = posTxt
            };
            var res = ed.Drag(jig);
            if (res.Status == PromptStatus.OK)
            {
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    btr.AppendEntity(jig.MLeader());
                    tr.AddNewlyCreatedDBObject(jig.MLeader(), true);
                    tr.TransactionManager.QueueForGraphicsFlush();
                    tr.Commit();
                }
            }
        }
        // Обработка выпадающего меню
        static void contextMenu_Popup(object sender, EventArgs e)
        {

            var contextMenu = sender as ContextMenuExtension;

            if (contextMenu != null)
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                // This is the "Root context menu" item
                var rootItem = contextMenu.MenuItems[0];
                try
                {
                    var acSsPrompt = ed.SelectImplied();
                    var mVisible = true;
                    if (acSsPrompt.Status == PromptStatus.OK)
                    {
                        var set = acSsPrompt.Value;
                        var ids = set.GetObjectIds();
                        if (acSsPrompt.Value.Count == 1)
                        {
                            using (var tr = doc.TransactionManager.StartTransaction())
                            {
                                var entity = tr.GetObject(ids[0], OpenMode.ForRead) as Entity;
                                var mpProductToSave = XDataHelpersForProducts.NewFromEntity(entity) as MpProductToSave;
                                mVisible = mpProductToSave != null;
                            }
                        }
                        else mVisible = false;
                    }

                    rootItem.Visible = mVisible;

                }
                catch 
                {
                    //
                }
            }
        }
    }
    class MLeaderJig : DrawJig
    {
        private MLeader _mleader;
        public Point3d FirstPoint;
        public string MlText;
        private Point3d _secondPoint;
        private Point3d _prevPoint;

        public MLeader MLeader()
        {
            return _mleader;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            var jpo = new JigPromptPointOptions
            {
                BasePoint = FirstPoint,
                UseBasePoint = true,
                UserInputControls = UserInputControls.Accept3dCoordinates |
                                    UserInputControls.GovernedByUCSDetect,
                Message = "\nТочка вставки: "
            };

            var res = prompts.AcquirePoint(jpo);
            _secondPoint = res.Value;
            if (res.Status != PromptStatus.OK) return SamplerStatus.Cancel;
            if (CursorHasMoved())
            {
                _prevPoint = _secondPoint;
                return SamplerStatus.OK;
            }
            return SamplerStatus.NoChange;
        }
        private bool CursorHasMoved()
        {
            return _secondPoint.DistanceTo(_prevPoint) > 1e-6;
        }
        protected override bool WorldDraw(WorldDraw draw)
        {
            var wg = draw.Geometry;
            if (wg != null)
            {
                const string arrowName = "_NONE";
                var arrId = MpCadHelpers.GetArrowObjectId(arrowName);

                var mtxt = new MText
                {
                    Contents = MlText,
                    Location = _secondPoint,
                    Annotative = AnnotativeStates.True
                };
                mtxt.SetDatabaseDefaults();

                _mleader = new MLeader();
                var ldNum = _mleader.AddLeader();
                _mleader.AddLeaderLine(ldNum);
                _mleader.ContentType = ContentType.MTextContent;
                _mleader.ArrowSymbolId = arrId;
                _mleader.MText = mtxt;
                _mleader.TextAlignmentType = TextAlignmentType.LeftAlignment;
                _mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine;
                _mleader.TextAngleType = TextAngleType.HorizontalAngle;
                _mleader.EnableAnnotationScale = true;
                _mleader.Annotative = AnnotativeStates.True;
                _mleader.AddFirstVertex(ldNum, FirstPoint);
                _mleader.AddLastVertex(ldNum, _secondPoint);
                _mleader.LeaderLineType = LeaderType.StraightLeader;
                _mleader.SetDatabaseDefaults();

                draw.Geometry.Draw(_mleader);
            }
            return true;
        }
    }
}
