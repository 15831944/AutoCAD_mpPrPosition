using System;
using System.Collections.Generic;
using ModPlusAPI.Interfaces;

namespace mpPrPosition
{
    public class Interface : IModPlusFunctionInterface
    {
        public SupportedProduct SupportedProduct => SupportedProduct.AutoCAD;
        public string Name => "mpPrPosition";
        public string AvailProductExternalVersion => "2015";
        public string FullClassName => string.Empty;
        public string AppFullClassName => string.Empty;
        public Guid AddInId => Guid.Empty;
        public string LName => "Позиция изделий";
        public string Description => "Функция добавляет/удаляет позицию в выбранные изделия и позволяет проставить маркировку позиции";
        public string Author => "Пекшев Александр aka Modis";
        public string Price => "0";
        public bool CanAddToRibbon => true;
        public string FullDescription => string.Empty;
        public string ToolTipHelpImage => string.Empty;
        public List<string> SubFunctionsNames => new List<string>();
        public List<string> SubFunctionsLames => new List<string>();
        public List<string> SubDescriptions => new List<string>();
        public List<string> SubFullDescriptions => new List<string>();
        public List<string> SubHelpImages => new List<string>();
        public List<string> SubClassNames => new List<string>();
    }
}
