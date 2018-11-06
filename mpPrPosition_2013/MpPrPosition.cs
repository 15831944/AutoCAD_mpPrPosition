namespace mpPrPosition
{
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using System;
    using System.Collections.Generic;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.GraphicsInterface;
    using Autodesk.AutoCAD.Windows;
    using Autodesk.AutoCAD.Runtime;
    using mpProductInt;
    using ModPlus;
    using ModPlus.Helpers;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using Exception = System.Exception;

    public class MpPrPosition : IExtensionApplication
    {
        private const string LangItem = "mpPrPosition";

        [CommandMethod("ModPlus", "mpPrPosition", CommandFlags.UsePickSet)]
        public void AddPositions()
        {
            Statistic.SendCommandStarting(new Interface());

            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            
            var opts = new PromptSelectionOptions();
            opts.Keywords.Add(Language.GetItem(LangItem, "h2"));
            var kws = opts.Keywords.GetDisplayString(true);
            opts.MessageForAdding = "\n" + Language.GetItem(LangItem, "h1") + ": " + kws;
            opts.KeywordInput += delegate (object sender, SelectionTextInputEventArgs e)
            {
                if (e.Input.Equals(Language.GetItem(LangItem, "h2")))
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
                    var elementsExist = new List<string>();
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
                                // Если еще не было
                                if (!elements.Contains(element))
                                {
                                    // Подсвечиваем
                                    ent.Highlight();
                                    // Запрос пользователю
                                    var pso = new PromptStringOptions(
                                        Language.GetItem(LangItem, "h3") + ": " + element + ". " + Language.GetItem(LangItem, "h4") + ": ")
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
                                // Если еще не было
                                if (!elementsExist.Contains(element))
                                {
                                    // Маркировка
                                    AddPositionMarker(ent, 0, product.Position, element, out markType);
                                    elements.Add(element);
                                    elementsExist.Add(element);
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
                    ExceptionBox.Show(ex);
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
                MessageForAdding = "\n" + Language.GetItem(LangItem, "h5") + ": "
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
                    ExceptionBox.Show(ex);
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
                var pko = new PromptKeywordOptions(
                    Language.GetItem(LangItem, "h6") + ": " + element + " " +
                    Language.GetItem(LangItem, "h7") + ": " + posTxt + ": [" +
                    Language.GetItem(LangItem, "h8") + "]", "Nothing Text Leader")
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
        private const string LangItem = "mpPrPosition";
        public static ContextMenuExtension MpPrPositionMarkCme;
        public static void Attach()
        {
            if (MpPrPositionMarkCme == null)
            {
                MpPrPositionMarkCme = new ContextMenuExtension();
                var miEnt = new MenuItem(Language.GetItem(LangItem, "h9"));
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
                        ExceptionBox.Show(ex);
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
            if (sender is ContextMenuExtension contextMenu)
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
                                mVisible = XDataHelpersForProducts.NewFromEntity(entity) is MpProductToSave;
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

    internal class MLeaderJig : DrawJig
    {
        private const string LangItem = "mpPrPosition";
        private MLeader _mLeader;
        public Point3d FirstPoint;
        public string MlText;
        private Point3d _secondPoint;
        private Point3d _prevPoint;

        public MLeader MLeader()
        {
            return _mLeader;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            var jpo = new JigPromptPointOptions
            {
                BasePoint = FirstPoint,
                UseBasePoint = true,
                UserInputControls = UserInputControls.Accept3dCoordinates |
                                    UserInputControls.GovernedByUCSDetect,
                Message = "\n" + Language.GetItem(LangItem, "h10") + ": "
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
                var arrId = AutocadHelpers.GetArrowObjectId(AutocadHelpers.StandardArrowhead._NONE);

                var mText = new MText
                {
                    Contents = MlText,
                    Location = _secondPoint,
                    Annotative = AnnotativeStates.True
                };
                mText.SetDatabaseDefaults();

                _mLeader = new MLeader();
                var ldNum = _mLeader.AddLeader();
                _mLeader.AddLeaderLine(ldNum);
                _mLeader.ContentType = ContentType.MTextContent;
                _mLeader.ArrowSymbolId = arrId;
                _mLeader.MText = mText;
                _mLeader.TextAlignmentType = TextAlignmentType.LeftAlignment;
                _mLeader.TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine;
                _mLeader.TextAngleType = TextAngleType.HorizontalAngle;
                _mLeader.EnableAnnotationScale = true;
                _mLeader.Annotative = AnnotativeStates.True;
                _mLeader.AddFirstVertex(ldNum, FirstPoint);
                _mLeader.AddLastVertex(ldNum, _secondPoint);
                _mLeader.LeaderLineType = LeaderType.StraightLeader;
                _mLeader.SetDatabaseDefaults();

                draw.Geometry.Draw(_mLeader);
            }
            return true;
        }
    }
}
