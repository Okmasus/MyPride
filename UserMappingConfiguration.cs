using CvParser.DataAccess.Models;
using CvParser.DataTransfer;
using CvParser.DataTransfer.User;
using CvParser.Extensions;
using Mapster;

namespace CvParser.MappingConfigurations
{
    public class UserMappingConfiguration : IRegister
    {
        public void Register(TypeAdapterConfig config)
        {
            config.NewConfig<User, OutputUserDTO>()
                .Map(dto => dto.Name, user => user.Profile.Name)
                .Compile();

            config
                .NewConfig<User, UserWithRoleDto>()
                .Map(user => user.Id, user => user.ID)
                .Map(user => user.Role, user => user.Role)
                .Map(user => user.Name, user => user.Profile.Name);

            config
                .NewConfig<UserWithRoleDto, UserAsRecruiterDto>()
                .Map(user => user.Id, user => user.Id)
                .Map(user => user.IsRecruiter, user => (user.Role & UserRole.Recruiter) == UserRole.Recruiter)
                .Map(user => user.Name, user => user.Name);

            config.NewConfig<User, OutputUserDTO>()
                .Map(user => user.ID, user => user.ID)
                .Map(user => user.Name, user => user.Profile.Name);
        }
    }
}