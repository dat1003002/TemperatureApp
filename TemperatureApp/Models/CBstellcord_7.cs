namespace TemperatureApp.Models
{
    public class CBstellcord_7
    {
        public int id { get; set; }

        public int CB_7_POLE { get; set; }       // Áp lực đá
        public bool CB_7_STATUS { get; set; }     // Trạng thái ON/OFF
        public int CB_7_CLAMP { get; set; }       // Áp lực bàn kẹp

        public DateTime CB_7_TIMES { get; set; }
    }
}
