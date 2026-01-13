using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Threading.Tasks;
using KoenZomers.OneDrive.Api;
using KoenZomers.OneDrive.Api.Entities;
using Newtonsoft.Json;

namespace KoenZomersKeePassOneDriveSync.Providers
{
    /// <summary>
    /// Helper class for making direct Microsoft Graph API calls
    /// Used as a workaround for deprecated OneDrive endpoints
    /// </summary>
    internal static class GraphApiHelper
    {
        private const string GraphApiBaseUrl = "https://graph.microsoft.com/v1.0";

        /// <summary>
        /// Retrieves items the user is following in their OneDrive for Business.
        /// This is used as a fallback if sharedWithMe isn't available.
        /// Note: Users need to "follow" items in OneDrive web interface for them to appear here.
        /// </summary>
        /// <param name="oneDriveApi">Authenticated OneDrive API instance</param>
        /// <returns>Collection of OneDriveItem objects representing followed items</returns>
        public static async Task<OneDriveItemCollection> GetFollowingItems(OneDriveApi oneDriveApi)
        {
            return await GetItemsFromGraph(oneDriveApi, "/me/drive/following");
        }

        /// <summary>
        /// Retrieves items shared with the current user using Microsoft Graph.
        /// </summary>
        /// <param name="oneDriveApi">Authenticated OneDrive API instance</param>
        /// <returns>Collection of OneDriveItem objects representing items shared with the user</returns>
        public static async Task<OneDriveItemCollection> GetSharedWithMeItems(OneDriveApi oneDriveApi)
        {
            return await GetItemsFromGraph(oneDriveApi, "/me/drive/sharedWithMe");
        }

        private static async Task<OneDriveItemCollection> GetItemsFromGraph(OneDriveApi oneDriveApi, string requestPath)
        {
            // Get the access token from the OneDriveApi instance
            var accessToken = oneDriveApi.AccessToken != null ? oneDriveApi.AccessToken.AccessToken : null;

            if (string.IsNullOrEmpty(accessToken))
            {
                throw new InvalidOperationException("No valid access token available. Please authenticate first.");
            }

            using (var httpClient = CreateGraphHttpClient(accessToken))
            {
                var items = new List<OneDriveItem>();
                var nextRequest = requestPath;

                while (!string.IsNullOrEmpty(nextRequest))
                {
                    var response = await httpClient.GetAsync(nextRequest);

                    if (!response.IsSuccessStatusCode)
                    {
                        var errorContent = await response.Content.ReadAsStringAsync();
                        throw new HttpRequestException(
                            string.Format("Graph API request failed with status {0}: {1}", response.StatusCode, errorContent));
                    }

                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    var result = JsonConvert.DeserializeObject<GraphCollectionResponse>(jsonResponse);

                    if (result != null && result.Value != null && result.Value.Length > 0)
                    {
                        items.AddRange(result.Value);
                    }

                    nextRequest = result != null ? result.NextLink : null;
                }

                return new OneDriveItemCollection
                {
                    Collection = items.Count > 0 ? items.ToArray() : new OneDriveItem[0]
                };
            }
        }

        /// <summary>
        /// Creates an HttpClient configured for Microsoft Graph API calls with proper authentication and proxy support
        /// </summary>
        /// <param name="accessToken">The bearer token for authentication</param>
        /// <returns>Configured HttpClient instance</returns>
        private static HttpClient CreateGraphHttpClient(string accessToken)
        {
            var proxySettings = Utilities.GetProxySettings();
            var proxyCredentials = Utilities.GetProxyCredentials();

            var httpClientHandler = new HttpClientHandler
            {
                UseDefaultCredentials = proxyCredentials == null,
                UseProxy = proxySettings != null,
                Proxy = proxySettings
            };

            if (proxyCredentials != null && httpClientHandler.Proxy != null)
            {
                httpClientHandler.Proxy.Credentials = proxyCredentials;
            }

            var httpClient = new HttpClient(httpClientHandler)
            {
                BaseAddress = new Uri(GraphApiBaseUrl)
            };

            httpClient.DefaultRequestHeaders.Accept.Clear();
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            // Add user agent similar to other parts of the plugin
            var assemblyVersion = Assembly.GetCallingAssembly().GetName().Version;
            httpClient.DefaultRequestHeaders.Add("User-Agent",
                string.Format("KoenZomers KeePass OneDriveSync v{0}.{1}.{2}.{3}",
                    assemblyVersion.Major, assemblyVersion.Minor, assemblyVersion.Build, assemblyVersion.Revision));

            return httpClient;
        }

        /// <summary>
        /// Response wrapper for Graph API collection responses (using Newtonsoft.Json)
        /// </summary>
        private class GraphCollectionResponse
        {
            [JsonProperty("value")]
            public OneDriveItem[] Value { get; set; }

            [JsonProperty("@odata.nextLink")]
            public string NextLink { get; set; }
        }
    }
}
