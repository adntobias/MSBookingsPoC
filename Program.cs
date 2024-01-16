// Create Graph Client 
// Set the required scopes for authentication
var scopes = new[] { "https://graph.microsoft.com/.default" };

var builder = new ConfigurationBuilder()
                 .SetBasePath(System.Environment.CurrentDirectory)
                 .AddJsonFile($"appsettings.json", true, true);

IConfiguration config = builder.Build();

// Values from app registration
var clientId = config["clientId"];
var tenantId = config["tenantId"];
var clientSecret = config["clientSecret"];

// Set up Azure.Identity
var options = new ClientSecretCredentialOptions
{
    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
};

// Create client credentials for authentication
var clientSecretCredential = new ClientSecretCredential(
    tenantId, clientId, clientSecret, options);

// Initialize Graph Service Client
var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

// Define some variables
string bID = "";
string currentDay = DateTime.Now.DayOfWeek.ToString();
string itSupportServiceId = "";
List<string> itServiceStaffIdsList = new List<string>();
List<User> itServiceStaffUsers = new List<User>();
List<Appointment> plannedAppointments = new List<Appointment>();
DateTime start = DateTime.MaxValue;
DateTime end = DateTime.MaxValue;
TimeSpan timeSlotInterval = TimeSpan.MaxValue;

// Get minimum information about the business
var businesses = await graphClient.Solutions.BookingBusinesses.GetAsync();
if (businesses is not null && businesses.Value is not null)
{
    foreach (var business in businesses.Value)
    {
        System.Console.WriteLine($"Business ID: {business.Id}");
        bID = business.Id;

        var thisBusiness = await graphClient.Solutions.BookingBusinesses[bID].GetAsync();
        if (thisBusiness is not null)
        {
            // Iterate through business hours for the current day
            foreach (var day in thisBusiness.BusinessHours)
            {
                if (day.Day.ToString().Equals(currentDay))
                {
                    // Display time slots for the current day
                    foreach (var slot in day.TimeSlots)
                    {
                        DateTime.TryParse(slot.StartTime.ToString(), out start);
                        DateTime.TryParse(slot.EndTime.ToString(), out end);

                        System.Console.WriteLine($"Hours: {start:t} - {end:t}");
                    }

                    break;
                }
            }

            timeSlotInterval = thisBusiness.SchedulingPolicy.TimeSlotInterval ?? TimeSpan.FromMinutes(15);
            
            System.Console.WriteLine($"     📅 Appointments every {timeSlotInterval.Hours}:{timeSlotInterval.Minutes}h");
        }
    }
}

System.Console.WriteLine("\n");

// Retrieve business services
var services = await graphClient.Solutions.BookingBusinesses[bID].Services.GetAsync();
if (services is not null && services.Value is not null)
{
    foreach (var service in services.Value)
    {
        System.Console.WriteLine($"✨ Service: {service.DisplayName}");
        if (service.DisplayName is not null && service.DisplayName.Equals("IT support"))
        {
            itSupportServiceId = service.Id;
        }

        foreach (var staffId in service.StaffMemberIds)
        {
            // Not the Graph.Users ID
            itServiceStaffIdsList.Add(staffId);
        }
    }
}

System.Console.WriteLine("\n");

// Retrieve business staff members
var staffMembers = await graphClient.Solutions.BookingBusinesses[bID].StaffMembers.GetAsync();
if (staffMembers is not null && staffMembers.Value is not null)
{
    foreach (BookingStaffMember staff in staffMembers.Value)
    {
        System.Console.WriteLine($"👨‍💻 Staff ID: {staff.Id}\n"+
                                 $"     DisplayName: {staff.DisplayName}\n"+
                                 $"     Email: {staff.EmailAddress}\n"+
                                 $"     UseBusinessHours: {staff.UseBusinessHours}");
        try
        {
            // Retrieve user information using Graph API
            User user = await graphClient
                                .Users[staff.EmailAddress]
                                .GetAsync(x =>
                                {
                                    x.QueryParameters.Select = new string[]
                                    { "id, displayName" };
                                });
            itServiceStaffUsers.Add(user);
        }
        catch (Exception e)
        {
            // Handle exception if user retrieval fails
        }
    }
}

System.Console.WriteLine("\n");

// Define a known customer for booking
BookingCustomer knownCustomer = new BookingCustomer();
knownCustomer.DisplayName = "Guy I Know";
knownCustomer.EmailAddress = "some@mail.com";

bool foundCustomer = false;

// Retrieve business customers
var customers = await graphClient.Solutions.BookingBusinesses[bID].Customers.GetAsync();
if (customers is not null && customers.Value is not null)
{
    foreach (BookingCustomer customer in customers.Value)
    {
        System.Console.WriteLine($"🤷‍♂️: {customer.Id}\n"+
                                 $"     Name: {customer.DisplayName}\n"+
                                 $"     Email: {customer.EmailAddress}");

        if (knownCustomer.EmailAddress.Equals(customer.EmailAddress))
        {
            knownCustomer = (BookingCustomer)customer;
            foundCustomer = true;
            System.Console.WriteLine("⬆️ This one I know ⚠️\n");
        }
    }

    // If the known customer is not found, create a new one
    if (foundCustomer is false)
    {
        var createdCustomer = await graphClient.Solutions.BookingBusinesses[bID].Customers.PostAsync(knownCustomer);
        if (createdCustomer is not null)
        {
            System.Console.WriteLine($"New Customer ID: {createdCustomer.Id}");
            knownCustomer = (BookingCustomer)createdCustomer;
        }
    }
}

// Retrieve the actual customer ID in base64
try
{
    var test = await graphClient.Solutions.BookingBusinesses[bID].Customers[knownCustomer.Id].GetAsync();

    if (test is not null)
    {
        System.Console.WriteLine($"🤷‍♂️: {((BookingCustomer)test).EmailAddress}");
    }
}
catch (Exception e) { /* Handle exception if retrieval fails */ }

System.Console.WriteLine("\n");

// Retrieve staff availability for the business
var requestBody = new Microsoft.Graph.Solutions.BookingBusinesses.Item.GetStaffAvailability.GetStaffAvailabilityPostRequestBody
{
    // Specify staff IDs and time range
    StaffIds = itServiceStaffIdsList,
    StartDateTime = new DateTimeTimeZone
    {
        DateTime = start.ToString("yyyy-MM-ddTHH:mm:ss"),
        TimeZone = "W. Europe Standard Time",
    },
    EndDateTime = new DateTimeTimeZone
    {
        DateTime = end.ToString("yyyy-MM-ddTHH:mm:ss"),
        TimeZone = "W. Europe Standard Time",
    },
};

// Get staff availability
var staffAvailability = await graphClient.Solutions.BookingBusinesses[bID].GetStaffAvailability.PostAsync(requestBody);
if (staffAvailability is not null && staffAvailability.Value is not null)
{
    foreach (StaffAvailabilityItem item in staffAvailability.Value)
    {
        System.Console.WriteLine(item.StaffId);

        if (item.AvailabilityItems is not null)
        {
            foreach (var times in item.AvailabilityItems)
            {
                System.Console.WriteLine($"     {((times.Status == BookingsAvailabilityStatus.Available) ? "✅":"⚠️") }{times.Status}\n"+
                                         $"         ⏰: {times.StartDateTime.DateTime:t} - {times.EndDateTime.DateTime:t}");
            }
        }
    }
}

System.Console.WriteLine("\n");

// Retrieve business calendar view
var calenderView = await graphClient.Solutions.BookingBusinesses[bID].CalendarView.GetAsync((requestConfiguration) =>
{
    requestConfiguration.QueryParameters.Start = start.ToString("yyyy-MM-ddTHH:mm:ssZ");
    requestConfiguration.QueryParameters.End = end.ToString("yyyy-MM-ddTHH:mm:ssZ");
});

if (calenderView is not null && calenderView.Value is not null)
{
    System.Console.WriteLine($"{calenderView.Value.Count()} appointments today");
    if (calenderView.Value.Count > 0)
    {
        foreach (var appointment in calenderView.Value)
        {
            plannedAppointments.Add(
                new Appointment(DateTime.Parse(appointment.StartDateTime.DateTime), DateTime.Parse(appointment.EndDateTime.DateTime))
            );
        }
    }
}

// Retrieve business appointments - all appointments, not only today - no filter user calenderView instead
var appointments = await graphClient.Solutions.BookingBusinesses[bID].Appointments.GetAsync();
if (appointments is not null && appointments.Value is not null)
{
    //System.Console.WriteLine(appointments.Value);
}

// Get available time slots
List<TimeSlot> availableTimeSlots = GetAvailableTimeSlots(start, end, timeSlotInterval, plannedAppointments);

System.Console.WriteLine("\n");

// Display available time slots
System.Console.WriteLine($"📅 {start.Date.ToShortDateString()}:");
foreach (var timeSlot in availableTimeSlots)
{
    System.Console.WriteLine($"     🕒: {timeSlot.StartTime.ToShortTimeString()} - {timeSlot.EndTime.ToShortTimeString()}");
}

// Create a new appointment for booking 
//business createAppointment
Random random = new Random();
TimeSlot randomTime = availableTimeSlots[random.Next(availableTimeSlots.Count)];
var appointmentRequestBody = new BookingAppointment
{
	OdataType = "#microsoft.graph.bookingAppointment",
	ServiceId = itSupportServiceId,
	ServiceName = "IT Service",
	ServiceNotes = "Customer requires punctual service.",
	StaffMemberIds = itServiceStaffIdsList,
	StartDateTime = new DateTimeTimeZone
	{
		OdataType = "#microsoft.graph.dateTimeTimeZone",
		DateTime = randomTime.StartTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffK"),
		TimeZone = "UTC",
	},
    EndDateTime = new DateTimeTimeZone
	{
		OdataType = "#microsoft.graph.dateTimeTimeZone",
		DateTime =  randomTime.EndTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffK"),
		TimeZone = "UTC",
	},
	MaximumAttendeesCount = 5,
	Customers = new List<BookingCustomerInformationBase>
	{
		new BookingCustomerInformation
		{
			OdataType = "#microsoft.graph.bookingCustomerInformation",
			CustomerId = knownCustomer.Id,
			Name = knownCustomer.DisplayName,
			EmailAddress = knownCustomer.EmailAddress
		},
	},
    IsLocationOnline = true,
	Price = 10d,
	PriceType = BookingPriceType.FixedPrice
};

var createAppointment = await graphClient.Solutions.BookingBusinesses[bID].Appointments.PostAsync(appointmentRequestBody);
if(createAppointment is not null){
    System.Console.WriteLine($"Your Appointment ID: {createAppointment.Id}");
}


// Method to calculate and return available time slots
static List<TimeSlot> GetAvailableTimeSlots(DateTime startTime, DateTime endTime, TimeSpan duration, List<Appointment> plannedAppointments)
{
    List<TimeSlot> availableTimeSlots = new List<TimeSlot>();

    DateTime currentSlotStart = DateTime.Now; // startTime;

    // Add appointments from UtcNow until the first planned appointment starts
    if (plannedAppointments.Count > 0)
    {
        DateTime firstPlannedStart = plannedAppointments[0].StartTime;

        while (currentSlotStart.Add(duration) <= firstPlannedStart && currentSlotStart <= endTime)
        {
            availableTimeSlots.Add(new TimeSlot(currentSlotStart, currentSlotStart.Add(duration)));
            currentSlotStart = currentSlotStart.Add(duration);
        }
    }

    // Add time slots between planned appointments
    foreach (var appointment in plannedAppointments)
    {
        if (currentSlotStart < appointment.StartTime)
        {
            DateTime currentSlotEnd = appointment.StartTime;

            while (currentSlotEnd.Add(duration) <= endTime)
            {
                availableTimeSlots.Add(new TimeSlot(currentSlotEnd, currentSlotEnd.Add(duration)));
                currentSlotEnd = currentSlotEnd.Add(duration);
            }

            currentSlotStart = appointment.EndTime;
        }
    }

    // Check for remaining available time after the last appointment
    while (currentSlotStart.Add(duration) <= endTime)
    {
        availableTimeSlots.Add(new TimeSlot(currentSlotStart, currentSlotStart.Add(duration)));
        currentSlotStart = currentSlotStart.Add(duration);
    }

    // Remove duplicates and already taken time slots
    availableTimeSlots = availableTimeSlots
        .Distinct(new TimeSlotEqualityComparer())
        .Where(slot => plannedAppointments.All(appointment => !IsOverlapping(slot, appointment)))
        .ToList();

    return availableTimeSlots;
}

// Method to check if a time slot overlaps with an appointment
static bool IsOverlapping(TimeSlot timeSlot, Appointment appointment)
{
    return timeSlot.StartTime < appointment.EndTime && timeSlot.EndTime > appointment.StartTime;
}

// Definition of Appointment class
class Appointment
{
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }

    public Appointment(DateTime startTime, DateTime endTime)
    {
        StartTime = startTime;
        EndTime = endTime;
    }
}

// Definition of TimeSlot class
class TimeSlot
{
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }

    public TimeSlot(DateTime startTime, DateTime endTime)
    {
        StartTime = startTime;
        EndTime = endTime;
    }
}
// Definition of TimeSlotEqualityComparer class for removing duplicates

class TimeSlotEqualityComparer : IEqualityComparer<TimeSlot>
{
    public bool Equals(TimeSlot x, TimeSlot y)
    {
        return x.StartTime == y.StartTime && x.EndTime == y.EndTime;
    }

    public int GetHashCode(TimeSlot obj)
    {
        return HashCode.Combine(obj.StartTime, obj.EndTime);
    }
}

