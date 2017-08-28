using mpPInterface;

namespace mpPrPosition
{
    public class Interface : IPluginInterface
    {
        public string Name => "mpPrPosition";
        public string AvailCad => "2013";
        public string LName => "Позиция изделий";
        public string Description => "Функция добавляет/удаляет позицию в выбранные изделия и позволяет проставить маркировку позиции";
        public string Author => "Пекшев Александр aka Modis";
        public string Price => "0";
    }
}
