namespace Chapter11.Model
{
    public class Project
    {
        public string Name { get; set; }

        public override string ToString()
        {
            return string.Format("Project: {0}", Name);
        }
    }
}
