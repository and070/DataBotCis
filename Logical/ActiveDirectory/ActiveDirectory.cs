using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;

namespace DataBotV5.Logical.ActiveDirectory
{
    /// <summary>
    /// Clase Logical encargada de active directory.
    /// </summary>
    class ActiveDirectory
    {
        /// <summary>
        /// Método para desactivar usuarios en el Active Directory
        /// </summary>
        /// <param name="user">Usuario o correo de la persona</param>
        /// <param name="domainUserLogin">Usuario admin que hara el cambio</param>
        /// <param name="domainUserPass">Contraseña del usuario admin</param>
        /// <returns></returns>
        public bool InactiveUser(string user, string domainUserLogin = null, string domainUserPass = null)
        {
            PrincipalContext AD;
            try
            {
                if (domainUserLogin == null && domainUserPass == null)
                    AD = new PrincipalContext(ContextType.Domain);
                else
                    AD = new PrincipalContext(ContextType.Domain, "dcpa01.gbm.net", domainUserLogin, domainUserPass);

                user = user.ToLower().Replace("@gbm.net", "");
                UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(AD, user);
                userPrincipal.Enabled = false;
                userPrincipal.Save();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public bool ExistAD(string newUser)
        {
            UserPrincipal result = GetUserPrincipal(newUser);
            try
            {
                if (result.Enabled == true)
                    return true;
                else
                    return false;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public Dictionary<string, string> GetAdData(string user)
        {
            UserPrincipal result = GetUserPrincipal(user);

            Dictionary<string, string> ret = new Dictionary<string, string>();
            ret.Add("Name", result.GivenName);
            ret.Add("LastName", result.Surname);
            ret.Add("FullName", result.DisplayName);

            return ret;

        }

        //Métodos Privados
        private UserPrincipal GetUserPrincipal(string user)
        {
            user = user.ToLower().Replace("@gbm.net", "");
            PrincipalContext AD = new PrincipalContext(ContextType.Domain);
            UserPrincipal u = new UserPrincipal(AD);
            u.SamAccountName = user;

            PrincipalSearcher search = new PrincipalSearcher(u);
            UserPrincipal result = (UserPrincipal)search.FindOne();
            search.Dispose();
            return result;
        }
    }
}
