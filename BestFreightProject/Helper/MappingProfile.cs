using AutoMapper;
using BestFreightProject.Dtos;
using BestFreightProject.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Helper
{
    public class MappingProfile : Profile
    {
        public MappingProfile() => CreateMap<Excel, ExcelCreateDto>().ReverseMap();
    }
}
