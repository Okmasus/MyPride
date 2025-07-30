using CvParser.API.Controllers.CV.Requests;
using CvParser.API.Controllers.CV.Response;
using CvParser.API.Services.BlockResumeParsing.Interfaces;
using CvParser.API.Services.Interfaces;
using CvParser.DataAccess.Models.Enums;
using CvParser.DataAccess.Repositories;
using CvParser.DataAccess.Repositories.Interfaces;
using CvParser.DataTransfer;
using CvParser.DataTransfer.ActionResult;
using CvParser.DataTransfer.Cv;
using CvParser.DataTransfer.Queries;
using CvParser.DataTransfer.ResumeBlockParsing;
using CvParser.DataTransfer.User;
using CvParser.DataTransfer.Watermarks;
using CvParser.Document.DTO;
using CvParser.Document.Enums;
using CvParser.Document.Implementation;
using CvParser.Document.Interfaces;
using CvParser.Document.ResumeBlockParsing.Interfaces;
using CvParser.Extensions;
using MapsterMapper;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using OneOf;
using OneOf.Types;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Claims;

namespace CvParser.API.Controllers.CV;

/// <summary>
/// Контроллер работы с резюме hh.ru.
/// </summary>
[Route("api/[controller]")]
[Authorize]
[ApiController]
public class CVController : ControllerBase
{
    private const string DocxContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private const string DocContentType = "application/msword";
    private readonly IResumeService _resumeService;
    private readonly IKeySkillService _keySkillService;
    private readonly IHhQueryService _hhQueryService;
    private readonly IDocumentParser[] _documentParsers;
    private readonly IUserService _userService;
    private readonly IStaffRateService _staffRateService;
    private readonly IResumeBlockParsingService _resumeBlockParsingService;

    public CVController(
        IResumeService resumeService,
        IHhQueryService hhQueryService,
        IServiceProvider serviceProvider,
        IUserService userService,
        IKeySkillService keySkillService,
        IStaffRateService staffRateService,
        IResumeBlockParsingService resumeBlockParsingService)
    {
        _resumeService = resumeService;
        _hhQueryService = hhQueryService;
        _userService = userService;
        _documentParsers = serviceProvider
            .GetServices<IDocumentParser>()
            .ToArray();
        _keySkillService = keySkillService;
        _staffRateService = staffRateService;
        _resumeBlockParsingService = resumeBlockParsingService;
    }

    /// <summary>
    /// Создаёт кандидата либо по ссылке либо по файлу.
    /// </summary>
    /// <param name="FileExtension">Расширения файла</param>
    /// <param name="StaffRate">Рейт кандидата</param>
    /// <param name="ResumeFile">Файл резюме</param>
    /// <param name="ResumeFileUrl">Ссылка на резюме</param>
    /// <param name="CompanyName">Название компании</param>
    /// <returns>Id кандидата.</returns>
    [HttpPost("[action]")]
    public async Task<ActionResult<long>> BlockParseResume(
            GoogleFilesTypes FileExtension,
            int StaffRate,
            IFormFile? ResumeFile,
            string? ResumeFileUrl,
            string? CompanyName,
            CancellationToken cancellationToken)
    {
        var resumeResult = await _resumeBlockParsingService.ParseResume(new
            ResumeBlockParsingDto(FileExtension, StaffRate, ResumeFile, ResumeFileUrl, CompanyName),
            cancellationToken);

        if (resumeResult.IsT1)
            return BadRequest(resumeResult.AsT1.Value);

        return Ok(resumeResult.AsT0.StaffId);
    }

    /// <summary>
    /// Обновить резюме на основе уже существующего.
    /// </summary>
    /// <param name="Id">Уникальный номер резюме.</param>
    /// <param name="dto">Поля, которые нужно обновить.</param>
    /// <returns></returns>
    [HttpPost("{Id:long}/Update")]
    public async Task<IActionResult> UpdateCVManually(long Id, UpdateCvDto dto)
    {
        await _resumeService.UpdateCvManually(Id, dto);

        return Ok();
    }

    /// <summary>
    /// Обновить опыт работы на основе уже существующего.
    /// </summary>
    /// <param name="Id">уникальный номер опыта работы.</param>
    /// <param name="dto">Поля, которые нужно обновить.</param>
    /// <returns></returns>
    [HttpPost("WorkExperience/{Id:long}/Update")]
    public async Task<ActionResult<OneOf<Success, Error<string>>>> UpdateWorkExperienceManually(long Id, UpdateWorkExperienceDto dto)
    {
        var result = await _resumeService.UpdateCvManually(Id, dto);

        return result.Match(
            success => Ok(ResponseResult.Success(success)),
            error => Ok(ResponseResult.Error(error.Value)));
    }

    /// <summary>
    /// Удаление опыта работы на основе уже существующего.
    /// </summary>
    /// <param name="id">уникальный номер опыта работы.</param>
    /// <returns></returns>
    [HttpPost("WorkExperience/{id:long}/Delete")]
    public async Task<IActionResult> DeleteWorkExperienceManually(long id)
    {
        await _resumeService.DeleteWorkExpirienceManually(id);
        return Ok();
    }

    /// <summary>
    /// Добавление опыта работы на основе уже существующего.
    /// </summary>
    /// <returns></returns>
    [HttpPost("{id:long}/WorkExperience/Add")]
    public async Task<ActionResult<OneOf<WorkExperienceDto, WorkExperienceErrorDto, NotFound, Error<string>>>> AddWorkExperienceManually(long id, AddWorkExperienceDto addWorkExperience)
    {
        var additionResult = await _resumeService.AddWorkExperienceManually(id,
            addWorkExperience);

        return additionResult.Match(
            ResponseResult.Success,
            workExperienceError => ResponseResult.Error("The error by the data of the work experience.", workExperienceError),
            notFound => ResponseResult.Error("Resume not found", notFound),
            error => ResponseResult.Error("Error", error.Value));
    }

    /// <summary>
    /// Прибавляем к продолжительности опыта работы +n%.
    /// </summary>
    /// <param name="id">ID опыта работы.</param>
    /// <param name="percent">Процент на который нужно увеличить.</param>
    /// <returns></returns>
    [HttpPost("WorkExperience/{id:long}/ChangeExperience/{percent:int}")]
    public async Task<IActionResult> ChangeExperienceDurationManually(long id, int percent)
    {
        var changeResult = await _resumeService.ChangeExperienceDurationManually(id, percent);

        return changeResult.Match(
            _ => ResponseResult.Success(),
            error => ResponseResult.Error(error.Value));
    }

    /// <summary>
    /// Обновить ключевой навык на основе уже существующего.
    /// </summary>
    /// <param name="id">уникальный номер ключегого навыка.</param>
    /// <param name="dto">Поля, которые нужно обновить.</param>
    /// <returns></returns>
    [HttpPost("keySkill/{Id:long}/Update")]
    public async Task<IActionResult> UpdateKeySkillManually(long Id, UpdateKeySkillDto dto)
    {
        await _keySkillService.UpdateKeySkillManually(Id, dto);

        return Ok();
    }

    /// <summary>
    /// Обновить ключеввые навыки с применением сортировки по порядку.
    /// </summary>
    /// <param name="dto">Поля, которые нужно обновить.</param>
    /// <returns></returns>
    [HttpPost("keySkill/UpdateOrder")]
    public async Task<ActionResult<IReadOnlyCollection<KeySkillDto>>> UpdateOrderOfKeySkillsManually(UpdateKeySkillsDto[] dto)
        => Ok(await _keySkillService.UpdateOrderOfKeySkillsManually(dto));

    /// <summary>
    /// Создание нескольких ключевых навыков.
    /// </summary>
    /// <param name="id">Уникальный номер ключевого навыка.</param>
    /// <returns></returns>
    [HttpPost("{resumeId:long}/keySkill/AddMany")]
    public async Task<ActionResult<List<KeySkillDto>>> CreateKeySkillsManually(long resumeId, string[] newKeySkills)
    {
        var keySkills = await _keySkillService.AddKeySkills(resumeId, newKeySkills);

        return keySkills.Match<ActionResult<List<KeySkillDto>>>(
            keySkillDto => keySkillDto,
            _ => BadRequest());
    }

    /// <summary>
    /// Создание одного ключевого навыка.
    /// </summary>
    /// <param name="resumeId">Уникальный номер резюме.</param>
    /// <returns></returns>
    [HttpPost("{resumeId:long}/keySkill/AddOne")]
    public async Task<ActionResult<KeySkillDto>> CreateKeySkillManually(long resumeId)
    {
        var result = await _keySkillService.AddKeySkill(resumeId);

        return result.Match<ActionResult<KeySkillDto>>(
            keySkillDto => keySkillDto,
            _ => BadRequest());
    }

    /// <summary>
    /// Удаление ключевого навыка.
    /// </summary>
    /// <param name="Id">Уникальный номер ключевого навыка.</param>
    /// <returns></returns>
    [HttpPost("keySkill/{Id:long}/Delete")]
    public async Task<IActionResult> DeleteKeySkillManually(long Id)
    {
        await _keySkillService.DeleteKeySkill(Id);

        return Ok();
    }

    /// <summary>
    /// Обновить оброзования на основе уже существующего.
    /// </summary>
    /// <param name="Id">уникальный номер резюме.</param>
    /// <param name="dto">Поля, которые нужно обновить.</param>
    /// <returns></returns>
    [HttpPost("Education/{Id:long}/Update")]
    public async Task<IActionResult> UpdateEducationManually(long Id, UpdateEducationDto dto)
    {
        await _resumeService.UpdateCvManually(Id, dto);

        return Ok();
    }

    /// <summary>
    /// Удаление образования.
    /// </summary>
    /// <param name="id">Идентификатор образования.</param>
    [HttpPost("Education/{id:long}/Delete")]
    public async Task<IActionResult> DeleteEducationManually(long id)
    {
        await _resumeService.DeleteEducationManually(id);
        return Ok();
    }

    /// <summary>
    /// Получение даты последней выгрузки CV.
    /// </summary>
    /// <returns>Дата последней выгрузки CV.</returns>
    [HttpGet("[action]")]
    public async Task<ActionResult<DateTimeOffset>> GetLastCVDateTime()
    {
        var dateTime = await _resumeService.GetLastResumeDateTime();
        return Ok(dateTime);
    }

    /// <summary>
    /// Получение списка CV на основе запроса.
    /// </summary>
    /// <param name="filterParameters">Параметры фильтрации.</param>
    /// <param name="queryId">ID запроса.</param>
    /// <returns>Список CV из запроса.</returns>
    [HttpGet("{queryId:long}")]
    public async Task<ActionResult<List<EstimatedOutputResumeDto>>> GetResumes([FromQuery] ResumesFilterParameters filterParameters, long queryId)
        => Ok(await _resumeService.GetAllResumes(filterParameters, queryId));

    /// <summary>
    /// Получение списка CV на основе запроса.
    /// </summary>
    /// <param name="filterParameters">Параметры фильтрации.</param>
    /// <param name="queryId">ID запроса.</param>
    /// <returns>Список CV из запроса.</returns>
    [HttpGet("{queryId:long}/Streaming")]
    public async Task<ActionResult<List<EstimatedOutputResumeDto>>> GetResumesByStreaming(
        [FromQuery] ResumesFilterParameters filterParameters, long queryId)
        => Ok(await _resumeService.GetAllResumesByStreaming(filterParameters, queryId));

    /// <summary>
    /// Возвращает список ID неоцененных CV из запроса.
    /// </summary>
    /// <param name="queryId">ID запроса.</param>
    /// <returns>Список ID неоцененных CV.</returns>
    [HttpGet("{queryId:long}/UnratedIds")]
    public async Task<ActionResult<List<long>>> GetUnratedResumeIds(long queryId, [FromQuery] ResumesFilterParameters filterParameters)
    {
        var resumeIds = await _resumeService.GetUnratedResumesIds(filterParameters, queryId);

        return Ok(resumeIds);
    }

    /// <summary>
    /// Получение CV по ID.
    /// </summary>
    /// <param name="id">ID резюме.</param>
    /// <param name="cancellationToken"></param>
    /// <returns>Конкретное CV, найденное по ID.</returns>
    [HttpGet("resume/{id:long}")]
    public async Task<ActionResult<EstimatedOutputResumeDto>> GetResume(long id, CancellationToken cancellationToken)
        => Ok(await _resumeService.GetResumeAsync(id, cancellationToken));

    /// <summary>
    /// Получение количества релевантного опыта в годах
    /// </summary>
    /// <param name="id"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [HttpGet("[action]")]
    public async Task<ActionResult<double>> GetRelevantWorkExperienceOfResume(long id, CancellationToken cancellationToken)
    {
        var result = await _resumeService.GetRelevantWorkExperienceOfResume(id, cancellationToken);

        return Ok(result);
    }

    /// <summary>
    /// Оценивание CV.
    /// </summary>
    /// <param name="resumeId">ID резюме.</param>
    /// <param name="queryId">ID запроса.</param>
    /// <param name="isGood">True - понравилось, False - не понравилось.</param>
    /// <param name="cancellationToken"></param>
    [HttpPost("Rate")]
    public async Task<ActionResult<EstimatedOutputResumeDto>> RateResume(long resumeId, long queryId,
        bool isGood, CancellationToken cancellationToken)
    {
        var query = await _hhQueryService.GetHhQuery(queryId, cancellationToken);
        var result
            = await _resumeService.RateResume(resumeId, queryId, isGood, User.Identity!.Name!, query, cancellationToken);

        return result.Match(
            _ => Ok(result.Value),
            error => ResponseResult.Error("Not Found"));
    }

    /// <summary>
    /// Общая оценка резюме.
    /// </summary>
    /// <param name="id">Идентификатор резюме.</param>
    /// <param name="cancellationToken"></param>
    /// <returns>Общий рейтинг резюме.</returns>
    [HttpGet("AllRatingResume/{id:long}")]
    public async Task<ActionResult<double>> GetAllRatingResumeAsync(long id, CancellationToken cancellationToken)
    {
        var resume = await _resumeService.GetResumeAsync(id, cancellationToken);
        return resume is null || resume.Rating is null ? 0 : resume.Rating.RateAll.Value;
    }

    /// <summary>
    /// Изменение статуса CV.
    /// </summary>
    /// <param name="resumeId">Идентификатор резюме.</param>
    /// <param name="queryId">Идентификатор запроса.</param>
    /// <param name="status">Статус.</param>
    /// <param name="cancellationToken"></param>
    [HttpPost("Status")]
    public async Task<IActionResult> UpdateResumeStatus(long resumeId, long queryId, ResumeUserStatus status, CancellationToken cancellationToken)
    {
        var resume = await _resumeService.GetResumeAsync(resumeId, cancellationToken);
        var query = await _hhQueryService.GetHhQuery(queryId, cancellationToken);

        if (resume is null || query is null)
            return NotFound();

        await _resumeService.UpdateResume(queryId, resumeId, new UpdateResumeDto { Status = status }, User.FindFirst(ClaimTypes.NameIdentifier)?.Value, "Обновлен статус резюме");

        return Ok();
    }

    /// <summary>
    /// Добавление комментария к CV.
    /// </summary>
    /// <param name="resumeId">Идентификатор резюме.</param>
    /// <param name="queryId">Идентификатор запроса.</param>
    /// <param name="comment">Комментарий.</param>
    /// <param name="cancellationToken"></param>
    [HttpPost("Comment")]
    public async Task<IActionResult> UpdateResumeComment(long resumeId, long queryId, string comment, CancellationToken cancellationToken)
    {
        var resume = await _resumeService.GetResumeAsync(resumeId, cancellationToken);
        var query = await _hhQueryService.GetHhQuery(queryId, cancellationToken);

        if (resume is null || query is null)
            return NotFound();

        await _resumeService.UpdateResume(queryId, resumeId, new UpdateResumeDto { Comment = comment }, User.FindFirst(ClaimTypes.NameIdentifier)?.Value, "Обновлен комментарий у резюме");

        return Ok();
    }

    /// <summary>
    /// Обрабатывает CV под нужды компании.
    /// </summary>
    /// <param name="file">Файл с CV в формате docx.</param>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    /// <returns></returns>
    [HttpPost("DocX")]
    public async Task<IActionResult> DocXProcessedResume(IFormFile file)
    {
        await using var ms = new MemoryStream();
        await file.CopyToAsync(ms);
        var result = await _resumeService.DocXProcessedResume(new(
            ms,
            file.ContentType switch
            {
                DocxContentType => ContentType.DocX,
                DocContentType => ContentType.Doc,
                _ => throw new ArgumentOutOfRangeException(nameof(file.ContentType)),
            }));

        return result.Match<IActionResult>(
            document => File(
                 document.ToArray(),
                contentType: "application/octet-stream",
                fileDownloadName: file.FileName),
            error => BadRequest(error));
    }

    /// <summary>
    /// Обрабатывает CV под нужды компании.
    /// </summary>
    /// <param name="request"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [HttpGet("DocX")]
    public async Task<IActionResult> DocXProcessedResume([FromQuery] DocXProcessedResumeRequest request, CancellationToken cancellationToken)
    {
        var result = await _resumeService.CreateDocument(new(request.CvId, Template.defalt, request.Format), cancellationToken);

        return result.Match<IActionResult>(
            document =>
            {
                var encodedFileName = WebUtility.UrlEncode(document.Name);
                Response.Headers["Content-Disposition"] = $"attachment; filename=\"{encodedFileName}\"";
                Response.Headers["Access-Control-Expose-Headers"] = "Content-Disposition";

                return File(
                    document.Document.ToArray(),
                    "application/octet-stream");
            },
            error => BadRequest(error));
    }

    /// <summary>
    /// Возвращает все форматы резюме в формате docx.
    /// </summary>
    /// <returns></returns>
    [HttpGet("[action]")]
    public ActionResult<DownloadingCvFormats[]> GetAllFormatsOfDocX() =>
        Ok(Enum
            .GetValues<DownloadingCvFormats>()
            .Select(value => new GetAllFormatsOfDocXResponse(
                (int)value,
                typeof(DownloadingCvFormats)
                .GetField(value.ToString())?
                .GetCustomAttribute<DescriptionAttribute>()?
                .Description)));

    /// <summary>
    /// Обрабатывает CV под нужды компании.
    /// </summary>
    /// <param name="request">Запрос</param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [HttpGet("{cvId:long}/DocX/Public")]
    public async Task<IActionResult> DocXProcessedResumePublic([FromQuery] DocXProcessedResumePublicRequest request, CancellationToken cancellationToken)
    {
        var result = await _resumeService.CreateDocument(new(request.CvId, Template.Publuc, request.Format), cancellationToken);

        return result.Match<IActionResult>(
            document => File(
                document.Document.ToArray(),
                "application/octet-stream",
                document.Name),
            error => BadRequest(error));
    }

    /// <summary>
    /// Создать публичное CV.
    /// </summary>
    /// <returns></returns>
    [HttpPost("Create/FromDocument/Public")]
    public async Task<ActionResult<EstimatedOutputResumeDto>> CreateCVFromDocument(IFormFile file, CancellationToken cancellationToken)
    {
        await using var ms = new MemoryStream();
        await file.CopyToAsync(ms);
        return Ok(await _resumeService.CreateCVFromDocument(new(
            ms,
            file.ContentType switch
            {
                DocxContentType => ContentType.DocX,
                DocContentType => ContentType.Doc,
                _ => throw new ArgumentOutOfRangeException(nameof(file.ContentType)),
            }), cancellationToken));
    }

    /// <summary>
    /// Получить CV по преобразованной ссылке.
    /// </summary>
    /// <param name="linkId">Уникальный номер ссылки.</param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [HttpGet("Public/{linkId:guid}")]
    public async Task<ActionResult<EstimatedOutputResumeDto>> GetCvFromConvertedLink(Guid linkId, CancellationToken cancellationToken)
    {
        var resume = await _resumeService.GetCvByConvertedLinkAsync(linkId, cancellationToken);

        return resume is null
            ? ResponseResult.Error("Not found")
            : ResponseResult.Success(resume);
    }

    /// <summary>
    /// Получить CV по преобразованной ссылке, в docx формате.
    /// </summary>
    /// <param name="linkId">Уникальный номер ссылки.</param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [HttpGet("Public/{linkId:guid}/DocX")]
    public async Task<IActionResult> GetCvDocumentFromConvertedLink(Guid linkId, CancellationToken cancellationToken)
    {
        var cvId = await _resumeService.GetCvIdByConvertedLinkAsync(linkId, cancellationToken);
        return cvId.Match<IActionResult>(
            success => File(
                success.Document.ToArray(),
                contentType: "application/octet-stream",
                fileDownloadName: success.Name),
            error => ResponseResult.Error(error.Value));
    }

    /// <summary>
    /// Статусы Сделок.
    /// </summary>
    /// <returns></returns>
    [HttpGet("[action]")]
    public IActionResult JobSearchStatus([FromServices] IMapper mapper)
    {
        var enumValues = Enum.GetValues<JobSearchStatus>().Order().ToList();
        var result = mapper.From(enumValues).AdaptToType<IEnumerable<EnumStructureEntryDto>>();
        return Ok(result);
    }

    /// <summary>
    /// Парсинг резюме из файла.
    /// </summary>
    /// <param name="file">Файл резюме.</param>
    /// <param name="cancellationToken"></param>
    /// <returns>Идентификатор нового кандидата.</returns>
    [HttpPost("ParseOrganizationsCandidateFromDocument")]
    public async ValueTask<ActionResult<OneOf<long, Error<string>>>> ParseOrganizationsCandidateFromDocument(
        IFormFile file,
        CancellationToken cancellationToken)
    {
        if (await _userService.GetUserAsync(User.Identity!.Name!, cancellationToken) is not { } user)
            return Ok(new Error<string>("A user not found."));

        var result = await _resumeService.ParseCandidateFromDocument(
            file,
            _documentParsers,
            new UserDataForResumeFileParseDTO(user.ID, user.Email),
            cancellationToken);


        return result.Match(
            success => Ok(success),
            error => Ok(error));
    }

    /// Скачивание файла из гугл-доков по ссылке (эндпоинт для тестирования скачивания гугл-файлов).

    /// </summary>
    /// <param name="url">Адрес файла.</param>
    /// <param name="fileType">Формат файла.</param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [HttpGet("AddResumeFromGoogle")]
    public async ValueTask<ActionResult<OneOf<FileResult, Error<string>>>> AddResumeFromGoogle(
        string companyName,
        string url,
        FinishedParserDTO dto,
        GoogleFilesTypes fileType,
        CancellationToken cancellationToken)
    {
        if (await _resumeService.CreateResume(new CreateResumeByGoogleUrlDto(companyName, url, fileType), dto, cancellationToken) is not
            { } resume)
            return BadRequest();

        return Ok();
    }

    /// </summary>
    /// <param name="url">Адрес файла.</param>
    /// <param name="fileType">Формат файла.</param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [HttpGet("UpdateResumeFromGoogle")]
    public async ValueTask<ActionResult<OneOf<FileResult, Error<string>>>> UpdateResumeFromGoogle(
        long resumeId,
        FinishedParserDTO dto,
        CancellationToken cancellationToken)
    {
        await _resumeService.UpdateResume(resumeId, dto, cancellationToken);

        return Ok();
    }

    /// <summary>
    /// Парсинг ссылок с резюме из гугл таблиц.
    /// </summary>
    /// <param name="url">Ссылка на источник.</param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [HttpGet("ParseGoogleSourceWithBenches")]
    public async ValueTask<ActionResult<string[]>> ParseGoogleSourceWithBenches(
        string url,
        CancellationToken cancellationToken) =>
        Ok(await _resumeService.ParseGoogleSourceWithBenches(
            url,
            cancellationToken));

    /// <summary>
    /// Обновить внутренний рейт всех кандидатов с бенча.
    /// </summary>
    /// <param name="cancellationToken">Токен отмены.</param>
    [HttpPost("[action]")]
    public async Task<ActionResult> UpdateRate(CancellationToken cancellationToken)
    {
        await _staffRateService.UpdateStaffRates(cancellationToken);
        return Ok();
    }

    /// <summary>
    /// Добавляет вотермарку для компании.
    /// </summary>
    /// <param name="file">Изображение.</param>
    /// <param name="OrganizationID">ID компании.</param>
    /// <returns></returns>
    [HttpPost("AddOrganizationWatermark")]
    public async Task<IActionResult> AddOrganizationWatermark(IFormFile file, long OrganizationID, CancellationToken cancellationToken)
    {
        var result = await _resumeService.AddWatermark(new AddWatermarkRequestDTO(file, OrganizationID), cancellationToken);

        return result.Match<IActionResult>(
        _ => Ok(),
        error => BadRequest(error));
    }

    /// <summary>
    /// Возвращает список всех названий компаний.
    /// </summary>
    [HttpGet("[action]")]
    public async Task<List<CompanyNamesDto>> GetAllCompanyNames(CancellationToken cancellationToken)
        => await _resumeService.GetAllCompanyNames(cancellationToken);

    /// <summary>
    /// Возвращает список всех названий компаний.
    /// </summary>
    [HttpGet("[action]")]
    public async ValueTask<ActionResult<string[]>> GetAllCompanyNamesOfOrganization(CancellationToken cancellationToken)
        => Ok(await _resumeService.GetAllCompanyNamesOfOrganization(cancellationToken));

    [HttpPost("[action]")]
    public async Task<ActionResult<IReadOnlyCollection<ResumeCloningDto>>> CloneResumes(
        ResumeCloningWithOrganizationIdRequest resumeCloningRequest,
        CancellationToken cancellationToken)
    {
        var cloningResult = await _resumeService.CloneResumes(resumeCloningRequest, cancellationToken);

        if (cloningResult.Match(ok => null, error => error) is { } error)
            return BadRequest(error);

        return Ok(cloningResult.AsT0);
    }

    /// <summary>
    /// Получить все навыки.
    /// </summary>
    [HttpGet("[action]")]
    public async Task<ActionResult<IReadOnlyCollection<KeySkillDto>>> GetAllKeySkills(int id )
    {
        var keySkills = await _keySkillService.GetAllKeySkillsAsync(id);

        if (keySkills is null || !keySkills.Any())
            return NotFound("Навыки не найдены.");

        return Ok(keySkills);
    }

    [HttpPost("[action]")]
    public async Task<ActionResult> UpdateBenchResumes(IReadOnlyCollection<UpdateBenchResumesDto> update, CancellationToken cancellationToken)
        => await _resumeService.UpdateResumeOnActualResume(update, cancellationToken) is not { } error 
            ? Ok() 
            : BadRequest(error);

       
}