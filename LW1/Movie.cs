using System.Collections.Generic;

namespace LW1
{
    class Movie
    {
        public long Id { get; set; }

        public string Color { get; set; }

        public string DirectiorName { get; set; }

        public double CriticsScore { get; set; }

        public double Duration { get; set; }

        public double DirectorFacebookLikes { get; set; }

        public double Actor1FacebookLikes { get; set; }

        public double Actor2FacebookLikes { get; set; }

        public double Actor3FacebookLikes { get; set; }

        public string Actor1Name { get; set; }

        public string Actor2Name { get; set; }

        public string Actor3Name { get; set; }

        public double Gross { get; set; }

        public List<string> Genre { get; set; }

        public string MovieTitle { get; set; }

        public int UsersVotesCount { get; set; }

        public int CastTotalFacebookLikes { get; set; }

        public int FaceNumInPoster { get; set; }

        public List<string> PlotKeywords { get; set; }

        public string MovieImdbLink { get; set; }

        public int UserViewsCount { get; set; }

        public string Language { get; set; }

        public string Country { get; set; }

        public string ContentRating { get; set; }

        public double Budget { get; set; }

        public int TitleYear { get; set; }

        public double ImdbScore { get; set; }

        public double RatioAspect { get; set; }

        public int MovieFacebookLikes { get; set; }

        public double Unnamed { get; set; }

        public int WarSymbTitle { get; set; }

        public int PointSymbTitle { get; set; }
    }
}
