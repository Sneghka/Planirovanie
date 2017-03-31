using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie.Objects
{
    public class LoginPasswordList
    {

        public static bool IsUserExistInSpravochinkByLogin(string userLogin, List<LoginPassword> usersList )
        {
            return usersList.Any(row => row.Login == userLogin);
        }

        public static LoginPassword GetUserObjByUserIdLogin(string userLogin, List<LoginPassword> usersList)
        {
            return (from user in usersList
                    where user.Login == userLogin
                    select user).First();
        }

        public static string GetUserEmailByLogin(string userLogin, List<LoginPassword> usersList)
        {
            var email = from user in usersList
                        where user.Login == userLogin
                        select user.Email;

            return string.Join("", email);
        }
    }
}
