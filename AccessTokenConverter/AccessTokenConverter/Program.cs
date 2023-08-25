using System;
using System.Diagnostics;
using System.IdentityModel.Tokens.Jwt;

class Program
{
    public static void Main(string[] args)
    {
        // this is a sample access token geneareted from jwt.io , replace it with real access token generated after the OTP validation
        string accessToken = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNTE2MjM5MDIyfQ.SflKxwRJSMeKKF2QT4fwpMeJf36POk6yJV_adQssw5c";
        string userEmail = GetEmailFromJwt(accessToken);

        if (!string.IsNullOrEmpty(userEmail))
        {
            Debug.WriteLine($"User's email: {userEmail}");
        }
        else
        {
            Debug.WriteLine("Failed to retrieve user's email.");
        }
    }

    public static string GetEmailFromJwt(string jwtToken)
    {            
        JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
        JwtSecurityToken jwt = tokenHandler.ReadJwtToken(jwtToken);

        // Extract the required data from the JWT token
        string userEmail = jwt.Payload["name"]?.ToString();
        return userEmail;
    }
}
