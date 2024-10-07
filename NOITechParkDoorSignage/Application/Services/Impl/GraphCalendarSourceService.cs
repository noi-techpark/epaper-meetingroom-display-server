// SPDX-FileCopyrightText: NOI Techpark <digital@noi.bz.it>
//
// SPDX-License-Identifier: AGPL-3.0-or-later

using Hangfire;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using NOITechParkDoorSignage.Application.BackgroundJobs;
using NOITechParkDoorSignage.Application.Services.Interfaces;
using NOITechParkDoorSignage.Infrastructure.Data.Interfaces;
using Event = NOITechParkDoorSignage.Domain.Room.Event;

namespace NOITechParkDoorSignage.Application.Services.Impl
{
    public class GraphCalendarSourceService : ICalendarSourceService
    {
        private readonly ILogger<MicrosoftGraphJob> _logger;
        private readonly IRoomRepository _roomRepository;
        private readonly GraphServiceClient _graphServiceClient;

        public GraphCalendarSourceService(
            ILogger<MicrosoftGraphJob> logger,
            IRoomRepository roomRepository,
            GraphServiceClient graphServiceClient
            )
        {
            _logger = logger;
            _roomRepository = roomRepository ?? throw new ArgumentNullException(typeof(IRoomRepository).Name); ;
            _graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(typeof(GraphServiceClient).Name);
        }

        public async Task SyncWithSource()
        {
            try
            {
                _logger.LogInformation("Refresh Office 365 Data");

                // Get all rooms from database
                var rooms = await _roomRepository.List();
                _logger.LogInformation($"Found {rooms.Count} rooms in database");

                foreach (var room in rooms)
                {
                    _logger.LogInformation("Room: '{name}, email: {email} (location: {location})", room.DisplayName, room.Email, room.Location);

                    // Using the Microsoft Graph Client check if the specific Calendar exists
                    var calendar = await _graphServiceClient.Users[room.Email].Calendar.GetAsync();
                    EventCollectionResponse eventCollections = await GetValidEvents(room);

                    if (eventCollections == null)
                    {
                        _logger.LogInformation("No events found for room {email}", room.Email);
                        continue;
                    }

                    // Loop the eventCollections and insert the calendar events into the repository
                    if (eventCollections.Value == null)
                    {
                        _logger.LogInformation("No events found for room {email}", room.Email);
                        continue;
                    }

                    _logger.LogWarning($"Found {eventCollections.Value.Count} events for room {room.Email}:\n{String.Join("\n", eventCollections.Value.Select((e, i) => $"{i + 1}) {e.Subject} {e.Start.ToDateTime().ToString("HH:mm")} UTC - {e.End.ToDateTime().ToString("HH:mm")} UTC"))}");

                    if (eventCollections.Value.Count == 0)
                    {
                        room.Events.Clear();
                        await _roomRepository.UnitOfWork.SaveChangesAsync();
                        return;
                    }

                    foreach (var office365Event in eventCollections.Value)
                    {
                        var calendarEvent = new Event
                        {
                            Id = office365Event.Id,
                            Title = office365Event.Subject,
                            Description = office365Event.BodyPreview,
                            IsPrivate = office365Event.Sensitivity == Sensitivity.Private,
                            StartDate = office365Event.Start.ToDateTime(), // Dates are saved in UTC on MS Graph (NB_ tymesettings on the server has to be correct!)
                            EndDate = office365Event.End.ToDateTime(), // Dates are saved in UTC on MS Graph
                            Location = office365Event.Location.DisplayName,
                            Organizer = office365Event.Organizer.EmailAddress.Address,
                            Attendees = office365Event.Attendees.Select(a => a.EmailAddress.Address).ToList(),
                            RoomId = room.Id
                        };

                        if (room.Events.Any(e => e.Id == calendarEvent.Id))
                        {
                            _logger.LogInformation("Event {id} already exists in database..update", calendarEvent.Id);

                            var existingEvent = room.Events.FirstOrDefault(e => e.Id == calendarEvent.Id);
                            if (existingEvent != null)
                            {
                                existingEvent.Title = calendarEvent.Title;
                                existingEvent.Description = calendarEvent.Description;
                                existingEvent.IsPrivate = office365Event.Sensitivity == Sensitivity.Private;
                                existingEvent.StartDate = calendarEvent.StartDate;
                                existingEvent.EndDate = calendarEvent.EndDate;
                                existingEvent.Location = calendarEvent.Location;
                                existingEvent.Organizer = calendarEvent.Organizer;
                                existingEvent.Attendees = calendarEvent.Attendees;
                            }
                        }
                        else
                        {
                            room.Events.Add(calendarEvent);
                        }

                        // Delete events not more in the calendar
                        var eventsToDelete = room.Events.Where(e => !eventCollections.Value.Any(oe => oe.Id == e.Id)).ToList();
                        foreach (var eventToDelete in eventsToDelete)
                        {
                            _logger.LogInformation("Event {id} not found in Office 365 calendar..delete", eventToDelete.Id);
                            room.Events.Remove(eventToDelete);
                        }
                    }

                    await _roomRepository.UnitOfWork.SaveChangesAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error while refreshing Office 365 data");
            }
        }

        private async Task<EventCollectionResponse> GetValidEvents(Domain.Room.Room room)
        {
            // Get all Events from Office 365 of a specific email by using the Graph client
            return await _graphServiceClient.Users[room.Email].CalendarView.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.StartDateTime =
                    DateTime.Now.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK");
                requestConfiguration.QueryParameters.EndDateTime =
                    DateTime.Now.Date.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ssK");
            });
        }

        public async Task<bool> AddEventByRoomEmail(Event evt, string roomEmail)
        {
            var room = await _roomRepository.Get(roomEmail);

            if (room == null)
            {
                _logger.LogError($"Room {roomEmail} not found");
                return false;
            }

            try
            {
                // Add an event in Microsoft Graph 
                var newEvent = new Microsoft.Graph.Models.Event
                {
                    Subject = evt.Title,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = evt.Description
                    },
                    Start = new DateTimeTimeZone
                    {
                        DateTime = evt.StartDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK"),
                        TimeZone = "UTC"
                    },
                    End = new DateTimeTimeZone
                    {
                        DateTime = evt.EndDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK"),
                        TimeZone = "UTC"
                    },
                    Location = new Location
                    {
                        DisplayName = evt.Location
                    },
                    Attendees = evt.Attendees.Select(a => new Attendee
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "iotdoorsignage@noi.bz.it",
                            Name = "IoT DoorSignage"
                        },
                        Type = AttendeeType.Required
                    }).ToList(),
                    CreatedDateTime = DateTime.Now,
                    TransactionId = Guid.NewGuid().ToString()
                };
                //newEvent.AdditionalData.Add("MadeBy", "DoorSignage");

                // log the current timezone
                _logger.LogInformation($"AddEventByRoomEmail - current timezone: {TimeZoneInfo.Local.DisplayName}");
                var office365NewEvent = await _graphServiceClient.Users[roomEmail].Calendar.Events.PostAsync(newEvent);

                RecurringJob.TriggerJob("MicrosoftGraphJob");

                return true;
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
            {
                _logger.LogError($"Office365 error: {ex.Error.Code}, {ex.Error.Message}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error while adding event to Office 365 calendar");
                return false;
            }
        }

        public async Task<bool> DeleteAllEventsAddedByTheLabel(string roomEmail)
        {
            var room = await _roomRepository.Get(roomEmail);

            if (room == null)
            {
                _logger.LogError("Room {email} not found", roomEmail);
                return false;
            }

            // Get all Events from Office 365 of a specific email by using the Graph client
            var eventCollections = await GetValidEvents(room);

            if (eventCollections?.Value == null)
            {
                _logger.LogInformation("No events found for room {email}", room.Email);
                return true;
            }

            var lastInsertedEvents = eventCollections.Value
                .Where(e => e.Attendees != null && e.Attendees.Any(a => a.EmailAddress?.Address == "iotdoorsignage@noi.bz.it"))
                .ToList();

            foreach (var office365Event in lastInsertedEvents)
            {
                await _graphServiceClient.Users[room.Email].Calendar.Events[office365Event.Id].DeleteAsync();
            }

            RecurringJob.TriggerJob("MicrosoftGraphJob");

            return true;
        }
    }
}
