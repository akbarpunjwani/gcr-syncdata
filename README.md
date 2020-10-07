# Google Classroom: Sync Data Utility
This command line utility is useful for all who are supporting the automation of Google Classroom administration. Using the Google APIs this utility reads all important objects & fields related to classroom and returns the well-structured C# DataSet.

## Use Case / Scenario:
A school having 3000 students in multiple campuses went online with Google Classroom. IT Professional managed to arrange GSuite for Education license and utilized the bulk upload features to setup the classes based on data provided by school operations. After couple of weeks, there is a need to analyze current status of teacher / student enrollment and classroom activities. This console application would be usefull in below manner:
1. Fetches the **list of classrooms** with all key details (setup parameter is provided to control how many classes needs to be fetched)
2. Fetches the **list of people in each class** (registered, invited - teachers / students)
3. Fetches the **list of user profiles** of people (unique i-e if teacher is registered to 3 classes it would give single profile entry)
4. Fetches the **list of announcement** posted in each google classrooms (setup parameter is provided to control how many maximum announcement per class required)
5. Fetches the **list of topics in classwork** in each google classrooms (setup parameter is provided to control how many maximum topics per class required)
6. Persists all the data into XML file written from DataSet

## Usage
Follow below steps:
1. Clone this repository and open the solution file
2. If required, run the below command on NuGet Package Manager Console
    ` PM> Install-Package Google.Apis.Classroom.v1 `
3. Use the utility method as follows:
    ` GCRSyncData.Start("gcradmin@yourdomain.org"); `
4. Above utility will persist the XML file in local directory, which can be read into any dataset for further use.

## Future Work
Further objects of Google APIs would be read into DataSet majorly with respect to the ClassWork submissions from Students, such that student assessment report could be compiled.

## We ❤️ Contributors Like You!
We’re eager to work with you, our user community, to improve these materials and develop new ones. Please check out [CLASSROOM API guide](https://developers.google.com/classroom) for more information on getting started.
