namespace TemperatureApp.Models
{
    public class CBstellcord_5
    {
        public int id { get; set; }

        public int CB_5_POLE { get; set; }       // Áp lực đá
        public bool CB_5_STATUS { get; set; }     // Trạng thái ON/OFF
        public int CB_5_CLAMP { get; set; }       // Áp lực bàn kẹp

        public DateTime CB_5_TIMES { get; set; }
    }
}
