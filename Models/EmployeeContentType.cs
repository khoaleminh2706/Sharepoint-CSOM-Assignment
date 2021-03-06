﻿using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;

namespace CreateSPSite.Models
{
    public class EmployeeContentType : AbstractContentType
    {
        public EmployeeContentType(ClientContext clientContext): base(clientContext)
        {
            Name = "Employee";
            FieldsList = new List<AbstractField>()
            {
                new SiteColumnField(clientContext) { InternalName = "FirstName" },
                new NewColumnField(clientContext) 
                { 
                    InternalName = "EmailAdd",
                    DisplayName = "Email Address",
                    XmlSchema = $"<Field ID='{Guid.NewGuid()}' Type='Text' Name='EmailAdd' StaticName='EmailAdd' DisplayName='Email Address' />" 
                },
                new NewColumnField(clientContext) 
                { 
                    InternalName = "ShortDesc",
                    DisplayName = "Short Description",
                    XmlSchema = $"<Field ID='{Guid.NewGuid()}' Type='Note' Name='ShortDesc' StaticName='ShortDesc' DisplayName='Short Description' NumLines='6' RichText='TRUE' RichTextMode='FullHtml' IsolateStyles='TRUE' Sortable='FALSE' />" 
                },
                new NewColumnField(clientContext) 
                { 
                    InternalName = "ProgrammingLanguages",
                    DisplayName = "Programming Languages",
                    XmlSchema = $"<Field ID='{Guid.NewGuid()}' Type='Choice' Name='ProgrammingLanguages' StaticName='ProgrammingLanguages' DisplayName='Programming Languages' Format='Dropdown' FillInChoice='FALSE'><CHOICES>" +
                    $"<CHOICE>C#</CHOICE>" +
                    $"<CHOICE>F#</CHOICE>" +
                    $"<CHOICE>Java</CHOICE>" +
                    $"<CHOICE>Jquery</CHOICE>" +
                    $"<CHOICE>AngularJS</CHOICE>" +
                    $"<CHOICE>Other</CHOICE>" +
                    $"</CHOICES></Field>" 
                }
            };
        }
    }
}
