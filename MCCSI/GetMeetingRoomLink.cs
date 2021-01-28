using System.IO;

namespace MCCSI
{
    class GetMeetingRoomLink
    {
        public string GetMeetingRoomLinkMethod()
        {
            //read file to string variable



            string dataFile = File.ReadAllText(Form1.HiddenFile);
            string[] dataArray = dataFile.Split('|');


            string meetingRoomLink = dataArray[2];

            return meetingRoomLink;

        }
    }
}
