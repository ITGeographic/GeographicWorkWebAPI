using System.Collections.Generic;

namespace GeographicDynamicWebAPI.Wrappers
{
    public class XexiliResult
    {
        public bool Success { get; set; }              // Whether duplicates exist or not
        public string? Message { get; set; }           // Message for the frontend
        public List<DuplicateQarsafariDTO> Data { get; set; } = new List<DuplicateQarsafariDTO>();
    }

    public class DuplicateQarsafariDTO
    {
        public int LiterId { get; set; }
        public int UniqId { get; set; }
        public int UniqIdOld { get; set; }
    }
}
