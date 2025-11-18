using AutoMapper;
using GeographicDynamic_DAL.DTOs;
using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Interface;
using GeographicDynamic_DAL.Models;
using GeographicDynamicWebAPI.Wrappers;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.Repository
{
    public class VarjisFartiRepository : IVarjisFarti
    {
        private readonly IMapper _mapper;

        public VarjisFartiRepository(IMapper mapper)
        {
            _mapper = mapper;
        }

        public Result<bool> SaveVarjisFarti(List<VarjisFartiDTO> varjisFartiDTO)
        {

            GeographicDynamicDbContext con = new GeographicDynamicDbContext();
            // List<VarjisFarti> existingData = con.VarjisFartis.Where(x => x.AreaNameId == ).ToList();

            var entitiesToUpdate = _mapper.Map<List<VarjisFartiDTO>, List<VarjisFarti>>(varjisFartiDTO);

            foreach (var dto in varjisFartiDTO)
            {
                // Find the corresponding entity in existingData based on some key
                var existingEntity = con.VarjisFartis.FirstOrDefault(x => x.Id == dto.Id && x.AreaNameId == dto.AreaNameId);

                if (existingEntity != null)
                {
                    //Dictionary dictionary = con.Dictionaries.FirstOrDefault(x => x.Id == dto.SaxeobaId);
                    //dictionary.Name = dto.Name;
                    _mapper.Map(dto, existingEntity);
                    con.SaveChanges();
                }
                else
                {
                    // Handle scenario where corresponding entity is not found
                }
            }

            List<VarjisFarti> varjisFartisToInsert = entitiesToUpdate
                .Where(x => !con.VarjisFartis.Any(m => m.Id == x.Id))
                .ToList();
            foreach (var item in varjisFartisToInsert)
            {
                //VarjisFarti varjisFarti = new VarjisFarti();
                //varjisFarti.VarjisFarti1 = item.VarjisFarti1;
                //varjisFarti.AreaNameId = item.AreaNameId;

                //var varjisFarti = _mapper.Map<VarjisFartiDTO, VarjisFarti>(item);

                con.VarjisFartis.Add(item);
                con.SaveChanges();
            }
            return new Result<bool>
            {
                Success = true,
                StatusCode = System.Net.HttpStatusCode.OK
            };
        }

        public Result<bool> SaveSaxeobebi(List<DictionaryDTO> dictionaryDTO)
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            var saxeobebiToUpdate = _mapper.Map<List<DictionaryDTO>, List<Dictionary>>(dictionaryDTO);
            foreach (var dto in dictionaryDTO)
            {
                // Find the corresponding entity in existingData based on some key
                var existingEntity = geographicDynamicDbContext.Dictionaries.FirstOrDefault(x => x.Id == dto.ID);

                if (existingEntity != null)
                {
                    //Dictionary dictionary = con.Dictionaries.FirstOrDefault(x => x.Id == dto.SaxeobaId);
                    //dictionary.Name = dto.Name;
                    _mapper.Map(dto, existingEntity);
                    geographicDynamicDbContext.SaveChanges();
                }
                else
                {
                    // Handle scenario where corresponding entity is not found
                }
            }

            List<Dictionary> SaxeobebiToinsert = saxeobebiToUpdate
                .Where(x => !geographicDynamicDbContext.Dictionaries.Any(m => m.Id == x.Id)).OrderBy(x => x.Name)
                .ToList();

            foreach (var item in SaxeobebiToinsert)
            {

                geographicDynamicDbContext.Dictionaries.Add(item);
                geographicDynamicDbContext.SaveChanges();
            }


            return new Result<bool>
            {
                Success = true,
                StatusCode = System.Net.HttpStatusCode.OK
            };
        }


        public Result<VarjisFartiDTO> DeleteVarjisfarti(VarjisFartiDTO varjisfartiDTO)
        {
            try
            {
                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
                if (varjisfartiDTO != null)
                {
                    if (varjisfartiDTO.Id != null && varjisfartiDTO.AreaNameId != null)
                    {
                        string constr = "Data Source=WIN-IK4QOCMD77O;Initial Catalog=Geographic_Dynamic_DB;User Id=sa;Password=123;Trusted_Connection=True;TrustServerCertificate=True;";
                        SqlCommand cmd = new SqlCommand();

                        try
                        {
                            using (SqlConnection connection = new SqlConnection(constr))
                            {
                                connection.Open();

                                string sqlCommandText = $"DELETE FROM VarjisFarti  WHERE ID = {varjisfartiDTO.Id} AND AreaNameID = {varjisfartiDTO.AreaNameId}";

                                if (!string.IsNullOrEmpty(sqlCommandText))
                                {
                                    using (SqlCommand sqlCommand = new SqlCommand(sqlCommandText, connection))
                                    {
                                        sqlCommand.ExecuteNonQuery();
                                    }
                                }
                                else
                                {
                                    connection.Close();
                                }

                            }
                        }
                        catch (Exception ex) { Console.WriteLine("An error occurred: " + ex.Message); }

                    }
                }
                return new Result<VarjisFartiDTO>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {
                return new Result<VarjisFartiDTO>
                {
                    Success = false,
                    Data = null,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }



        }
    }
}
