using System.Collections.Generic;
using testdocker;

namespace OpenXmlPocDocker
{
    public class UserDetailsData
    {
        public static readonly List<UserDetails> Persons = new List<UserDetails>()
           {
               new UserDetails() {ID="1001", Name="ABCD", City ="City1", Country="Fance"},
               new UserDetails() {ID="1002", Name="PQRS", City ="City2", Country="UK"},
               new UserDetails() {ID="1003", Name="XYZZ", City ="City3", Country="US"},
               new UserDetails() {ID="1004", Name="LMNO", City ="City4", Country="UAE"},
          };
    }
}
