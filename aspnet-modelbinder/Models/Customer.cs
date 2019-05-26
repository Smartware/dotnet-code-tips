using Microsoft.AspNetCore.Http;

public class Customer
{
    public string LastName {get;set;}
    public string FirstName {get;set;}
    public IFormFile Passport {get;set;}
    public IFormFile Signature {get;set;}
}