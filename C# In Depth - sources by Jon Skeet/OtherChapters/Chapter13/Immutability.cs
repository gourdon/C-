using System.ComponentModel;

namespace Chapter13
{
    [Description("Listing 13.07")]
    class Immutability
    {
        class Message
        {
            public Message(string from, string to,
                string body, string subject = null,
                byte[] attachment = null)
            {
                // Normal initialization code would go here
            }
        }

        static void Main()
        {
            Message message = new Message(
                from: "skeet@pobox.com",
                to: "csharp-in-depth-readers@everywhere.com",
                body: "I hope you like the second edition",
                subject: "A quick message"
            );
        }
    }
}
