using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.Contacts;

using LumenWorks.Framework.IO.Csv;

namespace GoogleContactsUpload
{
	class Program
	{
		static void Main(string[] args)
		{
			string username = "user1@gmail.com";
			string password = "password1";

			Uri feedUri = new Uri(ContactsQuery.CreateContactsUri(username));

			RequestSettings rs = new RequestSettings("teztech-GoogleContactsUpload-1", username, password);
			ContactsRequest cr = new ContactsRequest(rs);

			using (StreamReader contactsReader = new StreamReader(@"C:\Users\pj\Documents\Databases\ContactsMerged.csv"))
			{
				using (CsvReader contactsCsv = new LumenWorks.Framework.IO.Csv.CsvReader(contactsReader, true, ',', '"', '"', '\0', true))
				{
					while (contactsCsv.ReadNextRecord())
					{
						Contact contact = new Contact();
						contact.Title = string.Format("{0} {1}", contactsCsv["First Name"], contactsCsv["Last Name"]);

						if (contactsCsv["Email Address"] != "")
						{
							EMail workEmail = new EMail(contactsCsv["Email Address"]);
							workEmail.Primary = true;
							workEmail.Rel = ContactsRelationships.IsWork;
							contact.Emails.Add(workEmail);
						}

						if (contactsCsv["Home Email"] != "")
						{
							EMail homeEMail = new EMail(contactsCsv["Home Email"]);
							homeEMail.Rel = ContactsRelationships.IsHome;
							contact.Emails.Add(homeEMail);
						}

						if (contactsCsv["Work Phone"] != "")
						{
							PhoneNumber workPhone = new PhoneNumber(contactsCsv["Work Phone"]);
							workPhone.Rel = ContactsRelationships.IsWork;
							contact.Phonenumbers.Add(workPhone);
						}

						if (contactsCsv["Home Phone"] != "")
						{
							PhoneNumber homePhone = new PhoneNumber(contactsCsv["Home Phone"]);
							homePhone.Rel = ContactsRelationships.IsHome;
							contact.Phonenumbers.Add(homePhone);
						}

						if (contactsCsv["Mobile Phone"] != "")
						{
							PhoneNumber mobilePhone = new PhoneNumber(contactsCsv["Mobile Phone"]);
							mobilePhone.Primary = true;
							mobilePhone.Rel = ContactsRelationships.IsMobile;
							contact.Phonenumbers.Add(mobilePhone);
						}

						if (contactsCsv["Work Fax"] != "")
						{
							PhoneNumber faxPhone = new PhoneNumber(contactsCsv["Work Fax"]);
							faxPhone.Rel = ContactsRelationships.IsFax;
							contact.Phonenumbers.Add(faxPhone);
						}

						if (contactsCsv["Home Address"] != "")
						{
							PostalAddress homeAddress = new PostalAddress();
							homeAddress.Value = FormatAddress(false, contactsCsv["Home Address"], contactsCsv["Home Address 2"], contactsCsv["Home City"], contactsCsv["Home State"], contactsCsv["Home ZipCode"], contactsCsv["Home Country"]);
							homeAddress.Rel = ContactsRelationships.IsHome;
							contact.PostalAddresses.Add(homeAddress);
						}

						if (contactsCsv["Work Address"] != "")
						{
							PostalAddress workAddress = new PostalAddress();
							workAddress.Value = FormatAddress(false, contactsCsv["Work Address"], contactsCsv["Work Address 2"], contactsCsv["Work City"], contactsCsv["Work State"], contactsCsv["Work ZipCode"], contactsCsv["Work Country"]);
							workAddress.Rel = ContactsRelationships.IsWork;
							contact.PostalAddresses.Add(workAddress);
						}

						if (contactsCsv["Organization"] != "")
						{
							Organization organization = new Organization();
							organization.Label = contactsCsv["Organization"];
							organization.Title = contactsCsv["Job Title"];
							contact.Organizations.Add(organization);
						}

						StringBuilder notes = new StringBuilder();

						if (contactsCsv["Web Page 1"] != "")
							notes.AppendFormat("Website: {0}", contactsCsv["Web Page 1"]);

						notes.AppendFormat("{0}{1}", notes.Length > 0 ? "\r\n" : "", contactsCsv["Notes"]);

						contact.Content = notes.ToString();

						Contact createdContact = cr.Insert(feedUri, contact);
					}
				}
			}
		}

		static public string FormatAddress(bool useBR, string address1, string address2, string city, string state, string zip, string country)
		{
			StringBuilder address = new StringBuilder();
			string newLine = useBR ? "<br />" : "\r\n";

			if(address1 != "") address.AppendFormat("{0}{1}", address.Length > 0 ? newLine : "",   address1);
			if(address2 != "") address.AppendFormat("{0}{1}", address.Length > 0 ? newLine : "",   address2);

			StringBuilder cityStateZip = new StringBuilder();

			cityStateZip.Append(city);
			if(state != "")    cityStateZip.AppendFormat("{0}{1}", cityStateZip.Length > 0 ? ", "  : "", state);
			if(zip != "")      cityStateZip.AppendFormat("{0}{1}", cityStateZip.Length > 0 ? "  " : "", zip);

			if(cityStateZip.Length > 0) address.AppendFormat("{0}{1}", address.Length > 0 ? newLine : "",   cityStateZip.ToString());
			if(country != "")           address.AppendFormat("{0}{1}", address.Length > 0 ? newLine : "",   country);

			return address.ToString();
		}

	}
}
