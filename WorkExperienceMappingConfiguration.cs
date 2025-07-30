using CvParser.DataAccess.Models;
using CvParser.DataTransfer;
using CvParser.DataTransfer.Cv;
using Mapster;

namespace CvParser.MappingConfigurations;

public class WorkExperienceMappingConfiguration : IRegister
{
    public void Register(TypeAdapterConfig config)
    {
        config
            .NewConfig<UpdateWorkExperienceDto, WorkExperienceModel>()
            .IgnoreNullValues(true)
            .Compile();

        config.NewConfig<AddWorkExperienceDto, WorkExperienceModel>()
            .Map(model => model.IsPublic, _ => true)
            .Compile();
    }
}