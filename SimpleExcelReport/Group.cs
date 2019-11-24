namespace SimpleExcelReport
{
    public class Group<TRow>
    {
        public  ColumnBase<TRow>[] Columns { get; }
        public bool HasBorder { get; private set; }
        public string Heading { get; private set; }
        public bool HasHeading => !string.IsNullOrWhiteSpace(Heading);
        public Group(ColumnBase<TRow>[] columns)
        {
            Columns = columns;
        }

        public Group<TRow> Border(bool border = true)
        {
            HasBorder = true;

            return this;
        }

        public Group<TRow> SetHeading(string heading)
        {
            Heading = heading;

            return this;
        }
    }
}
