using CvParser.DataTransfer;
using CvParser.Extensions;
using Mapster;

namespace CvParser.MappingConfigurations;

public class EnumMappingConfigurations : IRegister
{
    public void Register(TypeAdapterConfig config)
    {
        config.NewConfig<Enum, EnumStructureEntryDto>()
            .Map(dto => dto.Id, @enum => Convert.ToInt32(@enum))
            .Map(dto => dto.Value, @enum => @enum.ToString())
            .Map(dto => dto.Description, @enum => @enum.GetDescription())
        .Compile();
    }
}