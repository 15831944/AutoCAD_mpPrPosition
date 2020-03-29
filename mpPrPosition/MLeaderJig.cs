namespace mpPrPosition
{
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.GraphicsInterface;
    using ModPlus.Helpers;
    using ModPlusAPI;

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
            if (res.Status != PromptStatus.OK)
                return SamplerStatus.Cancel;
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