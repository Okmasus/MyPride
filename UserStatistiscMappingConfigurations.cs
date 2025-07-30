using CvParser.DataAccess.Models;
using CvParser.DataTransfer;
using Mapster;

namespace CvParser.MappingConfigurations;

public class UserStatistiscMappingConfigurations : IRegister
{
    public void Register(TypeAdapterConfig config)
    {
        config.NewConfig<MessageStatistics, OutputMessageStatisticsDto>()
            .Map(dto => dto.AllCount, 
                model => model.FavoritesCount + model.TrashCount)
            .Compile();
        
        config.NewConfig<CvStatistics, OutputCvStatisticsOfUserDto>()
            .Map(dto => dto.AllCount,
                model => model.DislikesCount + model.LikesCount)
            .Compile();
        
        config.NewConfig<CallsToCandidatesStatistics, OutputCallsToCandidatesStatisticsDto>()
            .Map(dto => dto.AllCount,
                model => model.PendingCallsCount 
                         + model.SuccessfulCallsCount 
                         + model.UnsuccessfulCallsCount)
            .Compile();
        
        config.NewConfig<AddingCandidatesStatistics, OutputAddingCandidatesStatisticsDto>()
            .Map(dto => dto.AllCount, 
                model => model.AddedByLinkCount + model.AddedViaQueryCount)
            .Compile();
        
        config.NewConfig<DealStatusStatistics, OutputDealStatusStatisticsDto>()
            .Map(dto => dto.AllCount, model => model.Done + model.Waiting + model.AtWork)
            .Compile();
    }
}