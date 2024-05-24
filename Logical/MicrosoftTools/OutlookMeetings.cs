using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Logical.MicrosoftTools
{
    /// <summary>
    /// Clase Logical para agendar, actualizar y eliminar Meetings de Outlook
    /// </summary>
    class OutlookMeetings
    {
        /// <summary>
        /// ID Unico de la meetings, se usa para modificar, eliminar y extraccion de valores.
        /// </summary>
        public string MeetingUID { get; set; }
        /// <summary>
        /// Metodo que se utiliza para la creacion de Meetings de Outlook
        /// </summary>
        /// <param name="subject">Titulo del Meeting</param>
        /// <param name="body">Informacion del cuerpot del Meeting</param>
        /// <param name="meetingStart">Fecha y hora de inicio del Meeting</param>
        /// <param name="duration">Duracion en minutos del Meeting</param>
        /// <param name="atendee">Correo de la persona a la que se le agenda el Meeting</param>
        public void CreateMeeting(string subject, string body, DateTime meetingStart ,int duration, string atendee)
        {
            //Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
            //Microsoft.Office.Interop.Outlook.AppointmentItem agendaMeeting = (Microsoft.Office.Interop.Outlook.AppointmentItem)
            //outlookApplication.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.
            //olAppointmentItem);

            //if (agendaMeeting != null)
            //{
            //    agendaMeeting.MeetingStatus =
            //    Microsoft.Office.Interop.Outlook.OlMeetingStatus.olMeeting;
            //    //agendaMeeting.Location = "Conference Room";
            //    agendaMeeting.Subject = subject;
            //    agendaMeeting.Body = body;
            //    agendaMeeting.Start =meetingStart;
            //    agendaMeeting.ResponseRequested = false;
            //    agendaMeeting.Duration = duration;
            //    Microsoft.Office.Interop.Outlook.Recipient recipient =
            //        agendaMeeting.Recipients.Add(atendee);
            //    recipient.Type =
            //        (int)Microsoft.Office.Interop.Outlook.OlMeetingRecipientType.olRequired;
            //    ((Microsoft.Office.Interop.Outlook._AppointmentItem)agendaMeeting).Send();
            //    Meeting_UID = agendaMeeting.EntryID;              
            //}
        }
        /// <summary>
        /// Metodo que se utiliza para actualizar una Meeting existente con un UID definido
        /// </summary>
        /// <param name="uid">UID identificador unico del Meeting</param>
        /// <param name="subject">Cambio en el titulo del Meeting</param>
        /// <param name="body">Cambio en el cuerpo del Meeting</param>
        /// <param name="duration">Cambio en la duracion de minutos del Meeting</param>
        /// <param name="meetingStart">Cambio en la fecha y hora de inicio del Meeting</param>
        public void UpdateMeeting(string uid, string subject, string body, int duration,string email,[Optional] DateTime meetingStart)
        {
            //Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
            //Microsoft.Office.Interop.Outlook.MAPIFolder calendar =
            //outlookApplication.Session.GetDefaultFolder(
            // Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar);
            //Microsoft.Office.Interop.Outlook.Items calendarItems = calendar.Items;

            //try
            //{
            //    var reu = outlookApplication.Session.GetItemFromID(uid) as AppointmentItem;
            //    while (reu.Recipients.Count > 0)
            //    {
            //        reu.Recipients.Remove(1);
            //    }
            //    //email = "aeramirez@gbm.net";
            //    reu.Recipients.Add(email);
            //    reu.MeetingStatus = OlMeetingStatus.olMeeting;
            //    if (subject != "")
            //    {
            //        reu.Subject = subject;
            //    }
            //    if (body != "")
            //    {
            //        reu.Body = body;
            //    }
            //    if (meetingStart != null)
            //    {
            //        reu.Start = meetingStart;
            //    }
            //    if (duration != 0)
            //    {
            //        reu.Duration = duration;
            //    }
            //    reu.Save();
            //    reu.Send();
            //}
            //catch (System.Exception ex)
            //{


            //}
        }
        /// <summary>
        /// Metodo que se utiliza para eliminar una Meeting del caledario con un UID definido
        /// </summary>
        /// <param name="uid">UID indentificador unico del Meeting</param>
        public void DeleteMeeting(string uid)
        {
        //    Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
        //    Microsoft.Office.Interop.Outlook.MAPIFolder calendar =
        //    outlookApplication.Session.GetDefaultFolder(
        //     Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar);

        //    try
        //    {
        //        var reu = outlookApplication.Session.GetItemFromID(uid) as AppointmentItem;
        //        reu.MeetingStatus = OlMeetingStatus.olMeetingCanceled;
        //        reu.ForceUpdateToAllAttendees = true;
        //        reu.Save();
        //        reu.Send();
        //        reu.Delete();
        //    }
        //    catch (System.Exception ex)
        //    {

               
        //    }

        }
    }
}
