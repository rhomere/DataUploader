using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Uploader.Models;

namespace Uploader
{
    public class DbService
    {
        public string ConnectionString => "Server=tcp:cityserver.database.windows.net,1433;Initial Catalog=City;Persist Security Info=False;User ID=rhomere;Password=Gundam01;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";

        public void Add(Parcel parcel)
        {
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();
                SqlCommand command = new SqlCommand("spInsertAddress", connection, transaction);
                command.CommandType = System.Data.CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@MunicipalNumber", parcel.MunicipalNumber);
                command.Parameters.AddWithValue("@Owner", parcel.Owner);
                command.Parameters.AddWithValue("@Owner2", parcel.Owner2);
                command.Parameters.AddWithValue("@MailingAddressLine1", parcel.MailingAddressLine1);
                command.Parameters.AddWithValue("@MailingAddressLine2", parcel.MailingAddressLine2);
                command.Parameters.AddWithValue("@City", parcel.City);
                command.Parameters.AddWithValue("@State", parcel.State);
                command.Parameters.AddWithValue("@Zip", parcel.Zip);
                command.Parameters.AddWithValue("@SiteAddress", parcel.SiteAddress);
                command.Parameters.AddWithValue("@StreetNumber", parcel.StreetNumber);
                command.Parameters.AddWithValue("@StreetPrefix", parcel.StreetPrefix);
                command.Parameters.AddWithValue("@StreetName", parcel.StreetName);
                command.Parameters.AddWithValue("@StreetNumberSuffix", parcel.StreetNumberSuffix);
                command.Parameters.AddWithValue("@StreetSuffix", parcel.StreetSuffix);
                command.Parameters.AddWithValue("@CondoUnit", parcel.CondoUnit);
                command.Parameters.AddWithValue("@SiteCity", parcel.SiteCity);
                command.Parameters.AddWithValue("@SiteZip", parcel.SiteZip);

                var reader = command.ExecuteNonQuery();

                transaction.Commit();
                connection.Close();
            }
        }
    }
}
