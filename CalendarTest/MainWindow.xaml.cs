using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Google.Apis.Auth.OAuth2;
using System.Threading;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Util.Store;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;
using Google.Contacts;
using Google.GData.Client;

namespace CalendarTest
{
    public partial class MainWindow
    {
        public string StoreKey = "drzob-google_calendar";
        private UserCredential _credential;
        private const string ContactsScope = "https://www.google.com/m8/feeds";
        private const string ApplicationName = "DrZob";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void AddCalendarEventClick(object sender, RoutedEventArgs e)
        {
            const string calendarEventDescription = "Test event for DrZob";
            var timeFrom = DateTime.Today.AddHours(4);
            var timeTo = DateTime.Today.AddHours(6);
            var calendarEvent = CreateCalendarEvent(calendarEventDescription, timeFrom, timeTo);

            try
            {
                var credentials = GetCredentials(new[] { CalendarService.Scope.Calendar }, User.Text);

                var calendarService = new CalendarService(new BaseClientService.Initializer
                                                  {
                                                      HttpClientInitializer = credentials,
                                                      ApplicationName = ApplicationName
                                                  });

                var calendars = calendarService.CalendarList.List().Execute();
                var calendar = calendars.Items.SingleOrDefault(x => x.Summary == "DrZob");

                if (calendar == null)
                {
                    MessageBox.Show("The calendar 'DrZob' does not exist!");
                    return;
                }

                var calendarEventResponse = calendarService.Events.Insert(calendarEvent, calendar.Id).Execute();
            }
            catch (Exception ex)
            {
                Console.WriteLine("A Google Apps error occurred:");
                Console.WriteLine(ex.Message);
            }
        }

        private void AddContactClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var settings = Authorize(new[]
                {
                    ContactsScope
                });
                if (settings == null)
                    return;

                var contactRequest = new ContactsRequest(settings);
                var contacts = contactRequest.GetContacts();
                foreach (var contact in contacts.Entries)
                {
                    Console.WriteLine(contact.Name.FullName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("A Google Apps error occurred:");
                Console.WriteLine(ex.Message);
            }
        }

        private UserCredential GetCredentials(IEnumerable<string> scopes, string user = "")
        {
            var secrets = GetClientSecrects();

            var credPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), ".credentials/drzob_calendar", user);

            return GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, scopes, StoreKey, CancellationToken.None,
                new FileDataStore(credPath, true)).Result;
        }

        private RequestSettings Authorize(string[] scopes)
        {
            var secrets = GetClientSecrects();

            var credPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), ".credentials/drzob_calendar.json");

            _credential = GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, scopes, StoreKey, CancellationToken.None,
                new FileDataStore(credPath, true)).Result;

            var parameters = new OAuth2Parameters
                             {
                                 AccessToken = _credential.Token.AccessToken,
                                 RefreshToken = _credential.Token.RefreshToken
                             };

            return new RequestSettings(ApplicationName, parameters);
        }

        private ClientSecrets GetClientSecrects()
        {
            if (string.IsNullOrEmpty(ClientIdTextBox.Text) || string.IsNullOrEmpty(ClientSecretTextBox.Text))
            {
                MessageBox.Show("You must enter ClientId and ClientSecret first!");
                return null;
            }
            return new ClientSecrets
            {
                ClientId = ClientIdTextBox.Text,
                ClientSecret = ClientSecretTextBox.Text
            };
        }

        public bool HasStoredTokens()
        {
            var store = new FileDataStore("GoogleCalendar.Test");
            var t = store.GetAsync<TokenResponse>(StoreKey);
            t.Wait(TimeSpan.FromMilliseconds(100));
            return t.Result != null;
        }

        private static Event CreateCalendarEvent(string description, DateTime timeFrom, DateTime timeTo)
        {
            var start = new EventDateTime { DateTime = timeFrom };
            var end = new EventDateTime { DateTime = timeTo };

            return new Event { Start = start, End = end, Description = description, Summary = description };
        }
    }
}
