using System;
using System.Collections.Generic;

namespace Model
{
    public class SampleData
    {
        private readonly static List<Defect> defects;
        private readonly static List<User> users;
        private readonly static List<Project> projects;
        private readonly static List<NotificationSubscription> subscriptions;
        
        public static readonly DateTime Start = May(1);
        public static readonly DateTime End = May(31);

        public static IEnumerable<Defect> AllDefects
        {
            get { return defects; }
        }

        public static IEnumerable<User> AllUsers
        {
            get { return users; }
        }

        public static IEnumerable<Project> AllProjects
        {
            get { return projects; }
        }

        public static IEnumerable<NotificationSubscription> AllSubscriptions
        {
            get { return subscriptions; }
        }

        public static class Projects
        {
            public static readonly Project SkeetyMediaPlayer = new Project { Name="Skeety Media Player" };
            public static readonly Project SkeetyTalk = new Project { Name="Skeety Talk" };
            public static readonly Project SkeetyOffice = new Project { Name="Skeety Office" };
        }

        public static class Users
        {
            public static readonly User TesterTim = new User("Tim Trotter", UserType.Tester);
            public static readonly User TesterTara = new User("Tara Tutu", UserType.Tester);
            public static readonly User DeveloperDeborah = new User("Deborah Denton", UserType.Developer);
            public static readonly User DeveloperDarren = new User("Darren Dahlia", UserType.Developer);
            public static readonly User ManagerMary = new User("Mary Malcop", UserType.Manager);
            public static readonly User CustomerColin = new User("Colin Carton", UserType.Customer);
        }

        static SampleData()
        {
            projects = new List<Project>
            {
                Projects.SkeetyMediaPlayer,
                Projects.SkeetyTalk,
                Projects.SkeetyOffice
            };

            users = new List<User>
            {
                Users.TesterTim,
                Users.TesterTara,
                Users.DeveloperDeborah,
                Users.DeveloperDarren,
                Users.ManagerMary,
                Users.CustomerColin
            };

            subscriptions = new List<NotificationSubscription>
            {
                new NotificationSubscription { Project=Projects.SkeetyMediaPlayer, EmailAddress="media-bugs@skeetysoft.com" },
                new NotificationSubscription { Project=Projects.SkeetyTalk, EmailAddress="talk-bugs@skeetysoft.com" },
                new NotificationSubscription { Project=Projects.SkeetyOffice, EmailAddress="office-bugs@skeetysoft.com" },
                new NotificationSubscription { Project=Projects.SkeetyMediaPlayer, EmailAddress="theboss@skeetysoft.com"}
            };

            defects = new List<Defect>
            {
                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(1),
                    CreatedBy = Users.TesterTim,
                    Summary = "MP3 files crash system",
                    Severity = Severity.Showstopper,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Accepted,
                    LastModified = May(23)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(3),
                    CreatedBy = Users.DeveloperDeborah,
                    Summary = "Text is too big",
                    Severity = Severity.Trivial,
                    AssignedTo = null,
                    Status = Status.Closed,
                    LastModified = May(9)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(3),
                    CreatedBy = Users.CustomerColin,
                    Summary = "Sky is wrong shade of blue",
                    Severity = Severity.Minor,
                    AssignedTo = Users.TesterTara,
                    Status = Status.Fixed,
                    LastModified = May(19)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(4),
                    CreatedBy = Users.DeveloperDarren,
                    Summary = "Can't play files more than 200 bytes long",
                    Severity = Severity.Major,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Reopened,
                    LastModified = May(23)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(6),
                    CreatedBy = Users.TesterTim,
                    Summary = "Installation is slow",
                    Severity = Severity.Trivial,
                    AssignedTo = Users.TesterTim,
                    Status = Status.Fixed,
                    LastModified = May(15)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(7),
                    CreatedBy = Users.ManagerMary,
                    Summary = "DivX is choppy on Pentium 100",
                    Severity = Severity.Major,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Accepted,
                    LastModified = May(29)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(8),
                    CreatedBy = Users.DeveloperDeborah,
                    Summary = "Client acts as virus",
                    Severity = Severity.Showstopper,
                    AssignedTo = null,
                    Status = Status.Closed,
                    LastModified = May(10)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(8),
                    CreatedBy = Users.DeveloperDarren,
                    Summary = "Subtitles only work in Welsh",
                    Severity = Severity.Major,
                    AssignedTo = Users.TesterTim,
                    Status = Status.Fixed,
                    LastModified = May(23)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(9),
                    CreatedBy = Users.CustomerColin,
                    Summary = "Voice recognition is confused by background noise",
                    Severity = Severity.Minor,
                    AssignedTo = null,
                    Status = Status.Closed,
                    LastModified = May(15)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(9),
                    CreatedBy = Users.TesterTim,
                    Summary = "User interface should be more caramelly",
                    Severity = Severity.Trivial,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Created,
                    LastModified = May(9)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(10),
                    CreatedBy = Users.ManagerMary,
                    Summary = "Burning a CD makes the printer catch fire",
                    Severity = Severity.Showstopper,
                    AssignedTo = null,
                    Status = Status.Closed,
                    LastModified = May(29)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(10),
                    CreatedBy = Users.TesterTara,
                    Summary = "Peer to peer pairing passes parameters poorly",
                    Severity = Severity.Minor,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Accepted,
                    LastModified = May(12)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(11),
                    CreatedBy = Users.DeveloperDarren,
                    Summary = "Delay when sending message",
                    Severity = Severity.Minor,
                    AssignedTo = Users.TesterTara,
                    Status = Status.Fixed,
                    LastModified = May(20)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(11),
                    CreatedBy = Users.ManagerMary,
                    Summary = "Volume control needs to go to 11",
                    Severity = Severity.Minor,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Created,
                    LastModified = May(11)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(11),
                    CreatedBy = Users.CustomerColin,
                    Summary = "Splash screen fades too quickly",
                    Severity = Severity.Minor,
                    AssignedTo = Users.TesterTara,
                    Status = Status.Fixed,
                    LastModified = May(15)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(12),
                    CreatedBy = Users.DeveloperDeborah,
                    Summary = "Text box doesn't keep up with fast typing",
                    Severity = Severity.Major,
                    AssignedTo = Users.DeveloperDeborah,
                    Status = Status.Accepted,
                    LastModified = May(12)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(12),
                    CreatedBy = Users.DeveloperDarren,
                    Summary = "Password displayed in plain text",
                    Severity = Severity.Showstopper,
                    AssignedTo = null,
                    Status = Status.Closed,
                    LastModified = May(13)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(12),
                    CreatedBy = Users.TesterTim,
                    Summary = "Play button points the wrong way",
                    Severity = Severity.Major,
                    AssignedTo = Users.TesterTim,
                    Status = Status.Fixed,
                    LastModified = May(17)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(13),
                    CreatedBy = Users.CustomerColin,
                    Summary = "Wizard needed for CD burning",
                    Severity = Severity.Minor,
                    AssignedTo = Users.CustomerColin,
                    Status = Status.Fixed,
                    LastModified = May(20)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(13),
                    CreatedBy = Users.ManagerMary,
                    Summary = "Subtitles don't display during fast forward",
                    Severity = Severity.Trivial,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Accepted,
                    LastModified = May(14)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(13),
                    CreatedBy = Users.DeveloperDarren,
                    Summary = "Memory leak when watching Memento",
                    Severity = Severity.Trivial,
                    AssignedTo = Users.DeveloperDeborah,
                    Status = Status.Created,
                    LastModified = May(13)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(13),
                    CreatedBy = Users.DeveloperDeborah,
                    Summary = "Profile screen shows login count of -1",
                    Severity = Severity.Major,
                    AssignedTo = Users.DeveloperDeborah,
                    Status = Status.Accepted,
                    LastModified = May(20)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(13),
                    CreatedBy = Users.TesterTim,
                    Summary = "Server crashes under heavy load (3 users)",
                    Severity = Severity.Major,
                    AssignedTo = Users.DeveloperDeborah,
                    Status = Status.Accepted,
                    LastModified = May(17)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(15),
                    CreatedBy = Users.TesterTara,
                    Summary = "Unable to connect to any media server",
                    Severity = Severity.Showstopper,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Reopened,
                    LastModified = May(18)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(15),
                    CreatedBy = Users.DeveloperDeborah,
                    Summary = "UI turns black and white when playing old films",
                    Severity = Severity.Minor,
                    AssignedTo = Users.TesterTara,
                    Status = Status.Fixed,
                    LastModified = May(25)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(16),
                    CreatedBy = Users.ManagerMary,
                    Summary = "Password reset changes passwords for all users",
                    Severity = Severity.Showstopper,
                    AssignedTo = null,
                    Status = Status.Closed,
                    LastModified = May(18)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(17),
                    CreatedBy = Users.TesterTim,
                    Summary = "Modern music sounds rubbish",
                    Severity = Severity.Trivial,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Created,
                    LastModified = May(17)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(18),
                    CreatedBy = Users.TesterTim,
                    Summary = "Webcam makes me look bald",
                    Severity = Severity.Showstopper,
                    AssignedTo = Users.TesterTim,
                    Status = Status.Fixed,
                    LastModified = May(27)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(18),
                    CreatedBy = Users.CustomerColin,
                    Summary = "Sound is distorted when speakers are underwater",
                    Severity = Severity.Major,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Created,
                    LastModified = May(18)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(19),
                    CreatedBy = Users.DeveloperDarren,
                    Summary = "Japanese characters don't display properly",
                    Severity = Severity.Major,
                    AssignedTo = Users.DeveloperDeborah,
                    Status = Status.Accepted,
                    LastModified = May(23)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(20),
                    CreatedBy = Users.TesterTara,
                    Summary = "Video takes 100% of CPU",
                    Severity = Severity.Major,
                    AssignedTo = Users.DeveloperDeborah,
                    Status = Status.Accepted,
                    LastModified = May(22)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(22),
                    CreatedBy = Users.TesterTim,
                    Summary = "DVD Easter eggs unavailable",
                    Severity = Severity.Trivial,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Created,
                    LastModified = May(22)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(23),
                    CreatedBy = Users.ManagerMary,
                    Summary = "Transparency is high for menus to be readable",
                    Severity = Severity.Minor,
                    AssignedTo = Users.DeveloperDeborah,
                    Status = Status.Accepted,
                    LastModified = May(25)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(24),
                    CreatedBy = Users.CustomerColin,
                    Summary = "About box is missing version number",
                    Severity = Severity.Minor,
                    AssignedTo = Users.CustomerColin,
                    Status = Status.Fixed,
                    LastModified = May(29)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(25),
                    CreatedBy = Users.TesterTim,
                    Summary = "Logs record confidential conversations",
                    Severity = Severity.Major,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Reopened,
                    LastModified = May(30)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(27),
                    CreatedBy = Users.DeveloperDeborah,
                    Summary = "Profanity filter is too aggressive",
                    Severity = Severity.Minor,
                    AssignedTo = Users.TesterTara,
                    Status = Status.Fixed,
                    LastModified = May(29)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(27),
                    CreatedBy = Users.TesterTara,
                    Summary = "Full screen mode fails on dual monitors",
                    Severity = Severity.Minor,
                    AssignedTo = Users.DeveloperDeborah,
                    Status = Status.Created,
                    LastModified = May(27)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(28),
                    CreatedBy = Users.CustomerColin,
                    Summary = "Visualization hypnotises pets",
                    Severity = Severity.Minor,
                    AssignedTo = Users.DeveloperDeborah,
                    Status = Status.Accepted,
                    LastModified = May(29)
                },

                new Defect
                {
                    Project = Projects.SkeetyTalk,
                    Created = May(29),
                    CreatedBy = Users.ManagerMary,
                    Summary = "Resizing while typing loses input",
                    Severity = Severity.Trivial,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Created,
                    LastModified = May(29)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(30),
                    CreatedBy = Users.TesterTim,
                    Summary = "Network is saturated when playing WAV file",
                    Severity = Severity.Minor,
                    AssignedTo = Users.TesterTim,
                    Status = Status.Fixed,
                    LastModified = May(31)
                },

                new Defect
                {
                    Project = Projects.SkeetyMediaPlayer,
                    Created = May(31),
                    CreatedBy = Users.TesterTara,
                    Summary = "Media library tells user to keep the noise down",
                    Severity = Severity.Major,
                    AssignedTo = Users.DeveloperDarren,
                    Status = Status.Created,
                    LastModified = May(31)
                }
            };
        }

        public static void AssignIDs()
        {
            AssignIDs (SampleData.AllProjects,
                       (project, id) => project.ProjectID = id);
            AssignIDs (SampleData.AllUsers,
                       (user, id) => user.UserID = id);
            AssignIDs (SampleData.AllDefects,
                       (defect, id) => defect.ID = id);
            AssignIDs (SampleData.AllSubscriptions,
                       (subscription, id) => subscription.NotificationSubscriptionID = id);
        }

        static void AssignIDs<T>(IEnumerable<T> elements,
                                 Action<T,int> action)
        {
            int id = 1;
            foreach (T element in elements)
            {
                action(element, id++);
            }
        }

        public static DateTime May(int day)
        {
            return new DateTime(2010, 5, day);
        }
    }
}
