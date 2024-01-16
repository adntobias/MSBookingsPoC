# MSBookingsPoC
PoC for MS Bookings Graph Endpoints

Calling: 
- BookingBusinesses
- Services
- StaffMembers
- Customers
- GetStaffAvailability
- CalendarView
- Appointments

## Some Notes
- Staff IDs are not the same as User IDs
  - To get them filter for Email (not UPN)
- Customer IDs need to be base64 encoded
  - Currently no GUIDs are supported
- Staff Available Times respect Calender Entries in Outlook
  - maybe takes some time to display in the MS Booking Calender
-  To get a List of free Timeslots you must write the logic the other way around
   - Create a List of all possible Timeslots and cancel out the occupied spaces

## Used Permissions
|Permission|Type|
|--|--|
|Bookings.Read.All|Application|
|BookingsAppointment.ReadWrite.All|Application|

##### :construction::construction: Ignore null warnings it's only a PoC :construction::construction: