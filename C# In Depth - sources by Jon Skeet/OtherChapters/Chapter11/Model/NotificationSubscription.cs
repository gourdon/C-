namespace Chapter11.Model
{
    public class NotificationSubscription
    {
        /// <summary>
        /// Project for which this subscriber is notified
        /// </summary>
        public Project Project { get; set; }

        /// <summary>
        /// The address to send the notification to
        /// </summary>
        public string EmailAddress { get; set; }
    }
}
