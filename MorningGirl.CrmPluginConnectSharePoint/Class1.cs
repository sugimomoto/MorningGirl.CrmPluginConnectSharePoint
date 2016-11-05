using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;

namespace MorningGirl.CrmPluginConnectSharePoint
{
    public class CreateSharePointFolder : IPlugin
    {
        // Azure ADに登録したアプリケーションのClient ID
        private string clientId = "****";

        // 実行ユーザーのID・PW
        private string userName = "***@***.onmicrosoft.com";
        private string password = "****!";
        
        // 対象のSharePointサイトURLと対象のドキュメントライブラリ名
        private string sharePointUrl = "https://****.sharepoint.com/";
        private string targetDocumentFolder = "****";

        /// <summary>
        /// SharePoint Folder作成Plugin
        /// 取引先企業レコード作成時に対象のSharePintドキュメントライブラリにフォルダを作成
        /// </summary>
        /// <param name="serviceProvider"></param>
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                // ContextからInputParameterを取得
                var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

                // InputParameterからエンティティを取得し、フォルダ名となる、取引先企業名を取得
                var entity = context.InputParameters["Target"] as Entity;
                var createFolderName = entity["name"];
                
                // Azure ADのAccess Tokenを生成
                var azureAdConnect = new AzureAdConnect(
                    clientId,
                    userName,
                    password,
                    sharePointUrl
                    );

                var token = azureAdConnect.GetAccessToken();

                token.Wait();

                // SharePointのフォルダ作成用リクエスト作成
                using (var httpClient = new HttpClient())
                {

                    httpClient.DefaultRequestHeaders.Add("Accept", "application/json; odata=verbose");

                    // 取得したAccess Tokenを指定
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Result.access_token);

                    // リクエストを作成
                    var createReq = new HttpRequestMessage(HttpMethod.Post, sharePointUrl + "_api/web/folders");
                    createReq.Content = new StringContent(string.Format("{ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': '/{0}/{1}'}", targetDocumentFolder, createFolderName));
                    createReq.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");

                    // リクエストを実行
                    var result = httpClient.SendAsync(createReq);
                    result.Wait();

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }

    /// <summary>
    /// AzureADのoAuth認証によりトークンを取得するためのクラス
    /// </summary>
    public class AzureAdConnect
    {
        private string _clientId { get; set; }
        private string _userName { get; set; }
        private string _password { get; set; }
        private string _resourceId { get; set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public AzureAdConnect(string clientId, string userName, string password, string recourceId)
        {
            _clientId = clientId;
            _userName = userName;
            _password = password;
            _resourceId = recourceId;
        }

        /// <summary>
        /// grant_type=passwordでTokenを取得
        /// </summary>
        /// <returns></returns>
        public async Task<AzureAccessToken> GetAccessToken()
        {
            var token = new AzureAccessToken();

            string oauthUrl = string.Format("https://login.windows.net/common/oauth2/token");
            string reqBody = string.Format("grant_type=password&client_id={0}&username={1}&password={2}&resource={3}&scope=openid",
                Uri.EscapeDataString(_clientId),
                Uri.EscapeDataString(_userName),
                Uri.EscapeDataString(_password),
                Uri.EscapeDataString(this._resourceId));

            var client = new HttpClient();
            var content = new StringContent(reqBody);
            content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/x-www-form-urlencoded");
            using (HttpResponseMessage response = await client.PostAsync(oauthUrl, content))
            {
                if (response.IsSuccessStatusCode)
                {
                    var serializer = new DataContractJsonSerializer(typeof(AzureAccessToken));
                    var json = await response.Content.ReadAsStreamAsync();
                    token = (AzureAccessToken)serializer.ReadObject(json);
                }
            }

            return token;
        }
    }

    /// <summary>
    /// Azure AD Access Token格納クラス
    /// </summary>
    [DataContract]
    public class AzureAccessToken
    {
        [DataMember]
        public string access_token { get; set; }

        [DataMember]
        public string token_type { get; set; }

        [DataMember]
        public string expires_in { get; set; }

        [DataMember]
        public string expires_on { get; set; }

        [DataMember]
        public string resource { get; set; }
    }
}
