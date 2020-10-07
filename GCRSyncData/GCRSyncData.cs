using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;

using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Data.OleDb;
using System.Globalization;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Classroom.v1;
using Google.Apis.Classroom.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google;

namespace GCRSyncData
{
    class GCRSyncData
    {
        static string[] Scopes = { ClassroomService.Scope.ClassroomCourses
                                    , ClassroomService.Scope.ClassroomRosters
                                    , ClassroomService.Scope.ClassroomProfileEmails
                                    , ClassroomService.Scope.ClassroomAnnouncements
                                    , ClassroomService.Scope.ClassroomProfileEmails
                                    , ClassroomService.Scope.ClassroomTopics
                                };
        //static string[] Scopes = { ClassroomService.Scope.ClassroomCoursesReadonly };
        static string ApplicationName = "GCR Sync Data";

        public static void ConnectDB()
        {
            //TODO: Create specific user for AIMS and/or set permissions to restrict access to other AIMS table

            //DateTime maxDBTimeStamp = Convert.ToDateTime(DAL.ExecuteReader("SELECT GetDate() [Timestamp]").Tables[0].Rows[0]["Timestamp"]);
        }

        public static void Start(string gcrAdminEmail)
        {
            try
            {

                //Test Code Only
                UserCredential credential;

                using (var stream =
                    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = gcrAdminEmail.Split('@')[0] + gcrAdminEmail.Split('@')[1].Split('.')[0] + ".json";//"testtoken.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Classroom API service.
                var service = new ClassroomService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                int totalCourses = 1;
                int idxCourse = 0;
                int totalAnnPerCourse = 3;
                int idxAnn = 0;
                int totalTopicPerCourse = 3;
                int idxTopic = 0;

                DataSet dsGCRData = new DataSet("GCRData");
                dsGCRData.ReadXml(@"EmptyDS.xml");
                DataTable dtCourses = dsGCRData.Tables[0];
                DataTable dtPeople = dsGCRData.Tables[1];
                DataTable dtUserProfiles = dsGCRData.Tables[2];
                DataTable dtAnnouncements = dsGCRData.Tables[3];
                DataTable dtTopics = dsGCRData.Tables[4];
                /*
                DataTable dtCourses = DAL.ExecuteReader("SELECT * FROM GCR_Courses WHERE 1=2").Tables[0].Copy();
                dtCourses.TableName = "GCR_Courses";
                dsGCRData.Tables.Add(dtCourses);

                DataTable dtPeople = DAL.ExecuteReader("SELECT * FROM GCR_People WHERE 1=2").Tables[0].Copy();
                dtPeople.TableName = "GCR_People";
                dsGCRData.Tables.Add(dtPeople);

                DataTable dtUserProfiles = DAL.ExecuteReader("SELECT * FROM GCR_UserProfiles WHERE 1=1").Tables[0].Copy();
                dtUserProfiles.TableName = "GCR_UserProfiles";
                dsGCRData.Tables.Add(dtUserProfiles);

                DataTable dtAnnouncements = DAL.ExecuteReader("SELECT * FROM GCR_Announcements WHERE 1=2").Tables[0].Copy();
                dtAnnouncements.TableName = "GCR_Announcements";
                dsGCRData.Tables.Add(dtAnnouncements);

                DataTable dtTopics = DAL.ExecuteReader("SELECT * FROM GCR_Topics WHERE 1=2").Tables[0].Copy();
                dtTopics.TableName = "GCR_Topics";
                dsGCRData.Tables.Add(dtTopics);
                */
                DataRow drCurrCourse = null;
                DataRow drCurrPeople = null;
                DataRow drCurrUserProfile = null;
                DataRow drCurrAnnouncement = null;
                DataRow drCurrTopic = null;

                CoursesResource.ListRequest request = service.Courses.List();
                ListCoursesResponse response = request.Execute();
                if (response.Courses != null && response.Courses.Count > 0)
                {
                    foreach (var course in response.Courses)
                    {
                        if (idxCourse < totalCourses)
                        {
                            /*
                            if (DAL.ExecuteScalar("SELECT Count(*) FROM GCR_Courses WHERE CourseId='" + course.Id + "'") > 0)
                            {
                                Console.WriteLine(string.Format("Already Found ==> {0}: {1}", course.Id, course.Name));
                                continue;
                            }
                            */
                            idxCourse++;
                        }
                        else
                            break;

                        try
                        {
                            drCurrCourse = dtCourses.NewRow();
                            drCurrCourse["CourseId"] = course.Id;
                            drCurrCourse["CourseName"] = course.Name;
                            drCurrCourse["Section"] = course.Section;
                            drCurrCourse["Room"] = course.Room;
                            drCurrCourse["CourseDescHeading"] = course.DescriptionHeading;
                            drCurrCourse["CourseDesc"] = course.Description;
                            drCurrCourse["OwnerId"] = course.OwnerId;
                            drCurrCourse["AlternateLink"] = course.AlternateLink;
                            drCurrCourse["EnrollmentCode"] = course.EnrollmentCode;
                            drCurrCourse["CourseState"] = course.CourseState;
                            drCurrCourse["CalendarId"] = course.CalendarId;
                            drCurrCourse["CourseGroupEmail"] = course.CourseGroupEmail;
                            drCurrCourse["TeacherGroupEmail"] = course.TeacherGroupEmail;
                            drCurrCourse["TeacherFolder"] = course.TeacherFolder;
                            drCurrCourse["CourseMaterialSets"] = course.CourseMaterialSets;
                            drCurrCourse["CourseETags"] = course.ETag;
                            drCurrCourse["CourseCreationTime"] = course.CreationTime;
                            drCurrCourse["CourseUpdateTime"] = course.UpdateTime;

                            //dtCourses.Rows.Add(drCurrCourse);
                            //InsertDataRowToDB(drCurrCourse);
                            //Shifted the COURSE INSERT at the end

                            ListTeachersResponse teacherList = service.Courses.Teachers.List(course.Id).Execute();
                            if (teacherList.Teachers != null && teacherList.Teachers.Count > 0)
                            {
                                foreach (var ppl in teacherList.Teachers)
                                {
                                    drCurrPeople = dtPeople.NewRow();
                                    drCurrPeople["CourseId"] = ppl.CourseId;
                                    drCurrPeople["UserState"] = "REGISTERED";
                                    drCurrPeople["Role"] = "TEACHER";
                                    drCurrPeople["UserId"] = ppl.UserId;
                                    drCurrPeople["PeopleETag"] = ppl.ETag;

                                    //dtPeople.Rows.Add(drCurrPeople);
                                    InsertDataRowToDB(drCurrPeople);

                                    if (dtUserProfiles.Select("UserId='" + ppl.UserId + "'").Length == 0)
                                    {
                                        //New User Profile
                                        drCurrUserProfile = dtUserProfiles.NewRow();
                                        drCurrUserProfile["UserId"] = ppl.UserId;
                                        drCurrUserProfile["EmailAddress"] = ppl.Profile.EmailAddress;
                                        drCurrUserProfile["GivenName"] = ppl.Profile.Name.GivenName;
                                        drCurrUserProfile["FamilyName"] = ppl.Profile.Name.FamilyName;
                                        drCurrUserProfile["FullName"] = ppl.Profile.Name.FullName;
                                        drCurrUserProfile["ProfileETag"] = ppl.Profile.ETag;
                                        drCurrUserProfile["NameETag"] = ppl.Profile.Name.ETag;

                                        //dtUserProfiles.Rows.Add(drCurrUserProfile);
                                        InsertDataRowToDB(drCurrUserProfile);
                                    }
                                }
                            }

                            ListStudentsResponse studList = service.Courses.Students.List(course.Id).Execute();
                            if (studList.Students != null && studList.Students.Count > 0)
                            {
                                foreach (var ppl in studList.Students)
                                {
                                    drCurrPeople = dtPeople.NewRow();
                                    drCurrPeople["CourseId"] = ppl.CourseId;
                                    drCurrPeople["UserState"] = "REGISTERED";
                                    drCurrPeople["Role"] = "STUDENT";
                                    drCurrPeople["UserId"] = ppl.UserId;
                                    drCurrPeople["PeopleETag"] = ppl.ETag;
                                    drCurrPeople["StudentWorkFolder"] = ppl.StudentWorkFolder;

                                    //dtPeople.Rows.Add(drCurrPeople);
                                    InsertDataRowToDB(drCurrPeople);

                                    if (dtUserProfiles.Select("UserId='" + ppl.UserId + "'").Length == 0)
                                    {
                                        //New User Profile
                                        drCurrUserProfile = dtUserProfiles.NewRow();
                                        drCurrUserProfile["UserId"] = ppl.UserId;
                                        drCurrUserProfile["EmailAddress"] = ppl.Profile.EmailAddress;
                                        drCurrUserProfile["GivenName"] = ppl.Profile.Name.GivenName;
                                        drCurrUserProfile["FamilyName"] = ppl.Profile.Name.FamilyName;
                                        drCurrUserProfile["FullName"] = ppl.Profile.Name.FullName;
                                        drCurrUserProfile["ProfileETag"] = ppl.Profile.ETag;
                                        drCurrUserProfile["NameETag"] = ppl.Profile.Name.ETag;

                                        //dtUserProfiles.Rows.Add(drCurrUserProfile);
                                        InsertDataRowToDB(drCurrUserProfile);
                                    }
                                }
                            }


                            InvitationsResource.ListRequest invreq = service.Invitations.List();
                            invreq.CourseId = course.Id;
                            ListInvitationsResponse inviteList = invreq.Execute();
                            if (inviteList.Invitations != null && inviteList.Invitations.Count > 0)
                            {
                                foreach (var inv in inviteList.Invitations)
                                {
                                    drCurrPeople = dtPeople.NewRow();
                                    drCurrPeople["CourseId"] = inv.CourseId;
                                    drCurrPeople["UserState"] = "INVITED";
                                    drCurrPeople["Role"] = inv.Role;
                                    drCurrPeople["UserId"] = inv.UserId;
                                    drCurrPeople["PeopleETag"] = inv.ETag;

                                    //dtPeople.Rows.Add(drCurrPeople);
                                    InsertDataRowToDB(drCurrPeople);

                                    if (dtUserProfiles.Select("UserId='" + inv.UserId + "'").Length == 0)
                                    {
                                        //New User Profile
                                        try
                                        {
                                            UserProfile invUP = service.UserProfiles.Get(inv.UserId).Execute();

                                            drCurrUserProfile = dtUserProfiles.NewRow();
                                            drCurrUserProfile["UserId"] = inv.UserId;
                                            drCurrUserProfile["EmailAddress"] = invUP.EmailAddress;
                                            drCurrUserProfile["GivenName"] = invUP.Name.GivenName;
                                            drCurrUserProfile["FamilyName"] = invUP.Name.FamilyName;
                                            drCurrUserProfile["FullName"] = invUP.Name.FullName;
                                            drCurrUserProfile["ProfileETag"] = invUP.ETag;
                                            drCurrUserProfile["NameETag"] = invUP.Name.ETag;
                                        }
                                        catch(Exception upErr)
                                        {
                                            drCurrUserProfile = dtUserProfiles.NewRow();
                                            drCurrUserProfile["UserId"] = inv.UserId;
                                            drCurrUserProfile["FullName"] = upErr.Message;
                                        }
                                        //dtUserProfiles.Rows.Add(drCurrUserProfile);
                                        InsertDataRowToDB(drCurrUserProfile);
                                    }
                                }
                            }

                            //Announcements
                            ListAnnouncementsResponse annList = service.Courses.Announcements.List(course.Id).Execute();
                            idxAnn = 0;
                            if (annList.Announcements != null && annList.Announcements.Count > 0)
                            {
                                foreach (var ann in annList.Announcements)
                                {
                                    if (idxAnn < totalAnnPerCourse)
                                        idxAnn++;
                                    else
                                        break;

                                    drCurrAnnouncement = dtAnnouncements.NewRow();
                                    drCurrAnnouncement["AnnId"] = ann.Id;
                                    drCurrAnnouncement["CourseId"] = ann.CourseId;
                                    drCurrAnnouncement["CreatorUserId"] = ann.CreatorUserId;
                                    drCurrAnnouncement["AssigneeMode"] = ann.AssigneeMode;
                                    drCurrAnnouncement["AnnText"] = ann.Text.ToString().Substring(0, ((ann.Text.Length > 1000) ? 1000 : ann.Text.Length));
                                    drCurrAnnouncement["AnnState"] = ann.State;
                                    drCurrAnnouncement["AnnScheduledTime"] = ((ann.ScheduledTime == null) ? DBNull.Value : ann.ScheduledTime);
                                    drCurrAnnouncement["AnnCreationTime"] = ann.CreationTime;
                                    drCurrAnnouncement["AnnUpdateTime"] = ann.UpdateTime;

                                    //dtAnnouncements.Rows.Add(drCurrAnnouncement);
                                    InsertDataRowToDB(drCurrAnnouncement);
                                }
                            }

                            //Announcements
                            ListTopicResponse topicList = service.Courses.Topics.List(course.Id).Execute();
                            idxTopic = 0;
                            if (topicList.Topic != null && topicList.Topic.Count > 0)
                            {
                                foreach (var topic in topicList.Topic)
                                {
                                    if (idxTopic < totalTopicPerCourse)
                                        idxTopic++;
                                    else
                                        break;

                                    drCurrTopic = dtTopics.NewRow();
                                    drCurrTopic["TopicId"] = topic.TopicId;
                                    drCurrTopic["CourseId"] = topic.CourseId;
                                    drCurrTopic["TopicName"] = topic.Name;
                                    drCurrTopic["TopicETag"] = topic.ETag;
                                    drCurrTopic["TopicUpdateTime"] = topic.UpdateTime;

                                    //dtTopics.Rows.Add(drCurrTopic);
                                    InsertDataRowToDB(drCurrTopic);
                                }
                            }
                            drCurrCourse["TeacherCount"] = (teacherList.Teachers != null) ? teacherList.Teachers.Count.ToString() : "0";
                            drCurrCourse["StudentCount"] = (studList.Students != null) ? studList.Students.Count.ToString() : "0";
                            drCurrCourse["InviteCount"] = (inviteList.Invitations != null) ? inviteList.Invitations.Count.ToString() : "0";
                            drCurrCourse["AnnouncementCount"] = (annList.Announcements != null) ? annList.Announcements.Count.ToString() : "0";
                            drCurrCourse["TopicCount"] = (topicList.Topic != null) ? topicList.Topic.Count.ToString() : "0";

                            //dtCourses.Rows.Add(drCurrCourse);
                            InsertDataRowToDB(drCurrCourse);

                            Console.WriteLine();
                            Console.WriteLine(string.Format("{0}) {1} [TCH:{2}, STD:{3}, INV:{4}, ANN:{5}, TPC:{6}]"
                                , idxCourse.ToString()
                                , course.Name
                                , drCurrCourse["TeacherCount"]
                                , drCurrCourse["StudentCount"]
                                , drCurrCourse["InviteCount"]
                                , drCurrCourse["AnnouncementCount"]
                                , drCurrCourse["TopicCount"]
                                ));
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(string.Format("Exception occurred:{0}{1}", e.Message, e.StackTrace));
                            try
                            {
                                /*
                                DAL.ExecuteNonQuery(@"DELETE GCR_People WHERE CourseId NOT IN (SELECT CourseId FROM GCR_Courses)");
                                DAL.ExecuteNonQuery(@"DELETE GCR_Announcements WHERE CourseId NOT IN(SELECT CourseId FROM GCR_Courses)");
                                DAL.ExecuteNonQuery(@"DELETE GCR_Topics WHERE CourseId NOT IN(SELECT CourseId FROM GCR_Courses)");
                                */
                            }
                            catch { }
                            Console.WriteLine("Press Enter to continue with next course.");
                            Console.ReadLine();
                        }
                    }
                }
                Console.ReadLine();
                dsGCRData.WriteXml(@"DS_" + DateTime.Now.Ticks.ToString() + ".xml");
            }
            catch (Exception e)
            {
                Console.WriteLine(string.Format("Exception occurred:{0}{1}", e.Message, e.StackTrace));
                try
                {
                    /*
                    DAL.ExecuteNonQuery(@"DELETE GCR_People WHERE CourseId NOT IN (SELECT CourseId FROM GCR_Courses)");
                    DAL.ExecuteNonQuery(@"DELETE GCR_Announcements WHERE CourseId NOT IN(SELECT CourseId FROM GCR_Courses)");
                    DAL.ExecuteNonQuery(@"DELETE GCR_Topics WHERE CourseId NOT IN(SELECT CourseId FROM GCR_Courses)");
                    */
                }
                catch { }
                Console.ReadLine();
            }
        }

        protected static string currQueryTblName = "";
        protected static void InsertDataRowToDB(DataRow drGCRTableRow)
        {
            drGCRTableRow.Table.Rows.Add(drGCRTableRow);

            string insertSQLTemplate = @"
                                        INSERT INTO #TABLENAME#
                                                   ([CreatedOn]
		                                           #COLUMNNAME#
                                                   ,[ModifiedOn])
                                             VALUES
                                                   (GetDate()
		                                           #COLUMNVALUE#
                                                   ,GetDate())
                                        ";
            string query = insertSQLTemplate;
            foreach (DataColumn dc in drGCRTableRow.Table.Columns)
            {
                if (dc.ColumnName != "RecId" && dc.ColumnName != "CreatedOn" && dc.ColumnName != "ModifiedOn")
                {
                    query = query.Replace("#TABLENAME#", drGCRTableRow.Table.TableName)
                             .Replace("#COLUMNNAME#", string.Format(@",{0}
		                                           #COLUMNNAME#", dc.ColumnName))
                             .Replace("#COLUMNVALUE#", string.Format(@",'{0}'
		                                           #COLUMNVALUE#", drGCRTableRow[dc.ColumnName].ToString().Replace("'","''").Replace("’", "''").Replace("{", "(").Replace("}", ")")));
                }
            }
            query = query.Replace("#TABLENAME#", "").Replace("#COLUMNNAME#", "").Replace("#COLUMNVALUE#", "");
            //DAL.ExecuteNonQuery(query);
            Console.Write(".");

            if (currQueryTblName!= drGCRTableRow.Table.TableName)
            {
                //Console.WriteLine(query);
                //Console.ReadLine();
                currQueryTblName = drGCRTableRow.Table.TableName;
            }            
        }
    }
}
