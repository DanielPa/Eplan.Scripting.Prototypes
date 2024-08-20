using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net.Http;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace TranslateWithDeepl
{
  public class TranslateWithDeepl
  {
    private readonly string _apiKey;
    private readonly string _url = "https://api-free.deepl.com/v2/translate";
    private readonly HttpClient _httpClient;
    
    public async void Execute(string missingTermsFile, string translatedTermsFile)
    {
      var missingTerms = GetMissingTerms(missingTermsFile);
      await TranslateTerms(missingTerms, translatedTermsFile);
    }

    private async Task TranslateTerms(EplanLanguageDbRoot missingTerms, string translatedTermsFile)
    {
      var requestDto = new RequestDto();
      requestDto.TargetLang = "EN";
      requestDto.Text = missingTerms.TextSection.MT.SelectMany(mt => mt.T).Where(t => t.Lang == "de_DE").Select(t => t.Text).ToArray();
      var translate = await Translate(requestDto);
      
      for(int i = 0; i < missingTerms.TextSection.MT.Length; i++)
      {
        missingTerms.TextSection.MT[i].T.First(t => t.Lang == "en_US").Text = translate.Translations[i].Text;
      }

      var fileName = translatedTermsFile;
      var serializer = new XmlSerializer(typeof(EplanLanguageDbRoot));
      using (var writer = new StreamWriter(fileName))
      {
        serializer.Serialize(writer, missingTerms);
      }
      
      Process.Start(fileName);
    }

    public EplanLanguageDbRoot GetMissingTerms(string missingTermsFile)
    {
      var xmlFile = missingTermsFile;
      
      var serializer = new XmlSerializer(typeof(EplanLanguageDbRoot));
      EplanLanguageDbRoot eplanLanguageDbRoot = null;
      
      using (var reader = new StreamReader(xmlFile))
      {
        eplanLanguageDbRoot = (EplanLanguageDbRoot)serializer.Deserialize(reader);
      }
      return eplanLanguageDbRoot;
    }

    public TranslateWithDeepl()
    {
      // get api key from %DEEPL_API_KEY% environment variable
      _apiKey = System.Environment.GetEnvironmentVariable("DEEPL_API_KEY");
      _httpClient = new HttpClient();
    }

    private async Task<string> Translate(string text, string targetLang)
    {
      var translateDto = new RequestDto
      {
        Text = new string[] { text },
        TargetLang = targetLang
      };

      ResponseDto responseDto = await Translate(translateDto);
      return responseDto.Translations[0].Text;
    }

    private async Task<ResponseDto> Translate(RequestDto translateDto)
    {
      var json = JsonConvert.SerializeObject(translateDto);
      var data = new StringContent(json, Encoding.UTF8, "application/json");

      _httpClient.DefaultRequestHeaders.Add("Authorization", string.Format("DeepL-Auth-Key {0}", _apiKey));

      var response = await _httpClient.PostAsync(_url, data);
      var result = await response.Content.ReadAsStringAsync();
      ResponseDto responseDto = JsonConvert.DeserializeObject<ResponseDto>(result);

      return responseDto;
    }
  }

  public class ResponseDto
  {
    [JsonProperty("translations")]
    public TranslationDto[] Translations { get; set; }
  }

  public class TranslationDto
  {
    [JsonProperty("detected_source_language")]
    public string DetectedSourceLanguage { get; set; }

    [JsonProperty("text")]
    public string Text { get; set; }
  }

  public class RequestDto
  {
    [JsonProperty("text")]
    public string[] Text { get; set; }

    [JsonProperty("target_lang")]
    public string TargetLang { get; set; }
  }

  [XmlRoot(ElementName = "EplanLanguageDbRoot")]
  public class EplanLanguageDbRoot
  {
    [XmlAttribute(AttributeName = "Languages")]
    public string Languages { get; set; }

    [XmlAttribute(AttributeName = "numTexts")]
    public int NumTexts { get; set; }

    [XmlAttribute(AttributeName = "numTerms")]
    public int NumTerms { get; set; }

    [XmlAttribute(AttributeName = "numNonTranslate")]
    public int NumNonTranslate { get; set; }

    [XmlElement(ElementName = "TextSection")]
    public TextSection TextSection { get; set; }
  }

  public class TextSection
  {
    [XmlElement(ElementName = "MT")]
    public MT[] MT { get; set; }
  }

  public class MT
  {
    [XmlElement(ElementName = "T")]
    public T[] T { get; set; }
  }

  public class T
  {
    [XmlAttribute(AttributeName = "xml:lang")]
    public string Lang { get; set; }

    [XmlText]
    public string Text { get; set; }
  }
}