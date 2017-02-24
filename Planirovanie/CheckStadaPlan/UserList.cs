using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie.CheckStadaPlan
{
    public class UserList: List<User>
    {
        public static string GetUserEmailById(int userId, UserList emailSpravochick)
        {
            var email = from user in emailSpravochick
                where user.UserId == userId
                select user.Email;
            
            return string.Join("",email);
        }

       
    

    }
}
