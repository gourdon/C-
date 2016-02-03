using System.ComponentModel;

namespace Model
{
    public partial class User
    {
        public User (string name, UserType userType)
            : this()
        {
            Name = name;
            UserType = userType;
        }
        
        public override string ToString()
        {
            return string.Format("User: {0} ({1})", Name, UserType);
        }
    }
}
