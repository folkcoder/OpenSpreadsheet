namespace Tests.Models
{
    using System;

    public class DataTypes
    {
        public DataTypes()
        {
        }

        public DataTypes(bool loadPresets = false)
        {
            if (loadPresets)
            {
                this.Bool = true;
                this.Byte = 8;
                this.Char = 'z';
                this.Currency = -5423.34M;
                this.Decimal = -100.2323M;
                this.Double = 1.7E+3;
                this.Float = 10.2F;
                this.Int = -1212454;
                this.Long = 123423423423423;
                this.Text = "test string";
                this.DateTime = new DateTime(636826532750000000); // 2019-01-09 17:54:35
            }
        }

        public bool Bool { get; set; }

        public byte Byte { get; set; }

        public char Char { get; set; }

        public decimal Currency { get; set; }

        public DateTime DateTime { get; set; }

        public decimal Decimal { get; set; }

        public double Double { get; set; }

        public float Float { get; set; }

        public int Int { get; set; }

        public long Long { get; set; }

        public string Text { get; set; }
    }
}