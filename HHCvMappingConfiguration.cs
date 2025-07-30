using CvParser.DataAccess.Models;
using CvParser.DataTransfer;
using CvParser.HeadHunterAPI;
using CvParser.HeadHunterAPI.Dto.JSON;
using Mapster;
using Microsoft.IdentityModel.Tokens;
using Contact = CvParser.DataAccess.Models.Contact;
using Education = CvParser.DataAccess.Models.EducationsModels.Education;

namespace CvParser.MappingConfigurations;

public class HHCvMappingConfiguration : IRegister
{
    public void Register(TypeAdapterConfig config)
    {
        config
            .NewConfig<Primary1, Education>()
            .Map(education => education.Name, json => json.Name)
            .Map(education => education.Organization, json => json.Organization ?? string.Empty)
            .Map(education => education.Result, json => json.Result)
            .Map(education => education.Year, json => json.Year)
            .Compile();

        config
            .NewConfig<Experience1, WorkExperienceModel>()
            .Map(workExperience => workExperience.CompanyCity, json => json.area == null ? "-" : json.area.name)
            .Map(workExperience => workExperience.CompanyName, json => json.company)
            .Map(workExperience => workExperience.Description,
                json => string.IsNullOrEmpty(json.description) ? string.Empty : json.description)
            .Map(workExperience => workExperience.Duration, json => TimeSpan.Zero)
            .Map(workExperience => workExperience.Start, json => json.start)
            .Map(workExperience => workExperience.End, json => json.end)
            .Map(workExperience => workExperience.Position, json => json.position ?? "-")
            .Compile();

        config.NewConfig<Contact1, Contact>()
            .Map(contact => contact.Type, json => json.Type())
            .Map(contact => contact.value, json => json.Value())
            .Compile();

        config.NewConfig<HHResumeJson, ResumeModel>()
            .Map(resume => resume.About, json => json.skills)
            .Map(resume => resume.Age, json => json.age)
            .Map(resume => resume.City, json => json.area.name)
            .Map(resume => resume.Contacts, json => json.GetContacts())
            .Map(resume => resume.Country, json => string.Join(", ", json.citizenship.Select(x => x.name)), json => json.citizenship != null)
            .Map(resume => resume.DateOfBirth, json => json.birth_date)
            .Map(resume => resume.Education, json => json.education.level.name)
            .Map(resume => resume.Employment, json => string.Join(", ", json.employments.Select(x => x.name)), json => json.employments != null)
            .Map(resume => resume.GraduateYear, json => json.education.primary.Min())
            .Map(resume => resume.HHID, json => json.id)
            .Map(resume => resume.LastDownloadResume, json => DateTime.Now)
            .Map(resume => resume.LastUpdateTime, json => json.updated_at ?? json.created_at)
            .Map(resume => resume.Name,
                json => string.Join(' ', json.first_name, json.last_name, json.middle_name).Trim())
            .Map(resume => resume.ParseDateTime, json => DateTime.Now)
            .Map(resume => resume.Price, json => json.GetSalary())
            .Map(resume => resume.PriceCode, json => json.salary == null ? null : json.salary.currency)
            .Map(resume => resume.Schedule, json => string.Join(", ", json.schedules.Select(x => x.name)), json => json.schedules != null)
            .Map(resume => resume.SearchStatus, json => json.job_search_status.ToJobSearchStatus())
            .Map(resume => resume.Stack, json => json.title)
            .Map(resume => resume.URL, json => $"https://hh.ru/resume/{json.id}?hhtmFrom=resume_search_result")
            .Map(resume => resume.WorkExperience, json => json.experience)
            .Map(resume => resume.Educations, json => json.education.primary.Cast<IEducation>().Concat(json.education.additional))
            .Map(resume => resume.KeySkills, json => json.skill_set.Where(v => !string.IsNullOrEmpty(v)).Select((v, i) => new KeySkill { Value = v, Order = i }), json => json.skill_set != null)
            .Map(resume => resume.ContactLink, json => json.GetContactURL())
            .Map(resume => resume.CanGetContacts, json => json.paid_services.IsNullOrEmpty())
            .IgnoreNullValues(true)
            .Compile();

        config.NewConfig<UpdateHHQueryDto, HHQuery>()
            .IgnoreNullValues(true)
            .Map(dest => dest.UserID, src => src.UserId)
            .Compile();

        config.NewConfig<HHQuery, ReadHHQueryDto>()
            .MapWith(src => new(
                src.ID,
                src.Url,
                src.Name,
                src.Status.ToString(),
                src.HHQueryResumes.Count(x => x.IsGood != null && x.IsVisible),
                src.HHQueryResumes.Count(x => x.IsGood == true && x.IsVisible),
                src.HHQueryResumes.Count(hhr => hhr.IsVisible),
                src.LastParseTime,
                src.IsStreaming
            ))
            .Compile();
        config.NewConfig<UpdateResumeDto, HHQueryResume>()
            .IgnoreNullValues(true)
            .Compile();

        config.NewConfig<UpdateResumeDto, ResumeModel>()
            .IgnoreNullValues(true)
            .Compile();

        config.NewConfig<HeadHunterClientErrors, OutputHeadHunterStatisticsDTO>()
            .MapWith(hh => new(hh.DateTime, hh.Message))
            .Compile();
    }
}