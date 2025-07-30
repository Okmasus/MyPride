using CvParser.DataAccess.Extensions;
using CvParser.DataAccess.Models;
using CvParser.DataAccess.Models.Enums;
using CvParser.DataAccess.TgDatabase.Models;
using CvParser.DataTransfer;
using CvParser.Extensions;
using CvParser.HeadHunterAPI;
using CvParser.HeadHunterAPI.Dto.JSON;
using Mapster;

namespace CvParser.MappingConfigurations;

public class StaffMappingConfiguration : IRegister
{
    public void Register(TypeAdapterConfig config)
    {
        config
            .NewConfig<HHResumeJson, CreateStaffByCvDto>()
            .Map(dto => dto.FullName, json => json.GetFullName())
            .Map(dto => dto.Email, json => json.GetEmail())
            .Map(dto => dto.Phone, json => json.GetPhone())
            .Map(dto => dto.Contacts, json => json.GetContacts().Select(c => c.Value()))
            .Map(dto => dto.Location, json => json.area.name)
            .Map(dto => dto.Salary, json => json.GetSalary())
            .Map(dto => dto.PhotoUrl, json => json.photo.medium)
            .Map(dto => dto.Age, json => json.age)
            .Map(dto => dto.Birthday, json => json.birth_date)
            .CompileProjection();

        config
            .NewConfig<ResumeModel, CreateStaffByCvDto>()
            .Map(dto => dto.FullName, resume => resume.Name ?? string.Empty)
            .Map(dto => dto.Email, resume => resume.Contacts.FirstContact(ContactType.Email))
            .Map(dto => dto.Phone, resume => resume.Contacts.FirstContact(ContactType.Phone))
            .Map(dto => dto.Contacts, resume => resume.Contacts.GetContacts())
            .Map(dto => dto.Location, resume => resume.City)
            .Map(dto => dto.Salary, resume => resume.Price ?? 0)
            .Map(dto => dto.PhotoUrl, _ => string.Empty)
            .Map(dto => dto.Age, resume => resume.Age)
            .Map(dto => dto.Birthday, resume => resume.DateOfBirth)
            .Map(dto => dto.CvId, resume => resume.ID)
            .Map(dto => dto.StackContent, resume => string.Join('\n', resume.Stack, resume.About))
            .Map(dto => dto.RelevantWorkExperience, resume => TimeSpan.FromDays(resume.WorkExperience.ToDoubleRelevant() * 365.2425))
            .Map(dto => dto.WorkExperience, resume => TimeSpan.FromDays(resume.WorkExperience.ToDouble() * 365.2425))
            .Map(dto => dto.OrganizationId, resume => resume.OrganizationId)
            .Compile();

        config
           .NewConfig<StaffModel, StaffByCVIdDto>()
           .Map(dto => dto.Id, staff => staff.ID)
           .Map(dto => dto.Stack, staff => staff.Stack.Name)
           .Map(dto => dto.ExternalRate, staff => staff.ExternalRate)
           .Map(dto => dto.InternalRate, staff => staff.InternalRate)
           .Compile();
    }
}