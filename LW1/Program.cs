using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LW1
{
    class Program
    {
        static void Main(string[] args)
        {
            var movies = GetMoviesFromFile(@"F:\univer\DataAnalysis\DataAnalysis\LW1\bin\cinema1.xlsx");

            //Console.WriteLine(movies.Count());

            movies = ValidateList(movies);

            //Console.WriteLine(movies.Count());

            //1 Построить распределение жанров по рейтингу, прибыльности, лайкам
            //GroupGenreByRating(movies);
            //GroupGenreByGross(movies);
            //GroupGenreByLikes(movies);

            //2 Найти топ 20 связок (если такие есть) актер-режиссер, которые дают больше денег в прокате
            //GroupByActorDirectorIncome(movies);

            //3 Указать, фильмы из какой страны имеют лучший средний рейтинг. Проанализировать ответ и аргументировать вердикт
            //GroupByCountryAverageRating(movies);

            //4 Какой сюжет в среднем содержат фильмы жанра драма (plot_keywords). Есть ли у этих фильмов общий сюжетный ход. 
            //GroupByDramaKeyWords(movies);

            //5 Влияет ли возрастной рейтинг на бюджет фильма
            //GroupByRating(movies);

            //6 Построить модель регрессии, которая бы оценивала (predict), какой будет успех (финансовый) у нового фильма  Х
            PredictNewFilm(movies);

            Console.ReadLine();
        }

        private static List<Movie> GetMoviesFromFile(string filePath)
        {
            var movies = new List<Movie>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = xlPackage.Workbook.Worksheets.First();
                var totalRows = worksheet.Dimension.End.Row;
                var totalColumns = worksheet.Dimension.End.Column;

                for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                {
                    var idString = worksheet.Cells[rowNum, 1].Text;
                    var id = !String.IsNullOrEmpty(idString) ? Int64.Parse(idString) : 0;

                    var usersVotesCountString = worksheet.Cells[rowNum, 14].Text;
                    var usersVotesCount = !String.IsNullOrEmpty(usersVotesCountString) ? Int32.Parse(usersVotesCountString) : 0;

                    var castTotalFacebookLikesString = worksheet.Cells[rowNum, 15].Text;
                    var castTotalFacebookLikes = !String.IsNullOrEmpty(castTotalFacebookLikesString) ? Int32.Parse(castTotalFacebookLikesString) : 0;

                    var faceNumInPosterString = worksheet.Cells[rowNum, 17].Text;
                    var faceNumInPoster = !String.IsNullOrEmpty(faceNumInPosterString) ? Int32.Parse(faceNumInPosterString) : 0;

                    var userViewsCountString = worksheet.Cells[rowNum, 20].Text;
                    var userViewsCount = !String.IsNullOrEmpty(userViewsCountString) ? Int32.Parse(userViewsCountString) : 0;

                    var movieFacebookLikesString = worksheet.Cells[rowNum, 29].Text;
                    var movieFacebookLikes = !String.IsNullOrEmpty(movieFacebookLikesString) ? Int32.Parse(movieFacebookLikesString) : 0;

                    var warSymbTitleString = worksheet.Cells[rowNum, 31].Text;
                    var warSymbTitle = !String.IsNullOrEmpty(warSymbTitleString) ? Int32.Parse(warSymbTitleString) : 0;

                    var pointSymbTitleString = worksheet.Cells[rowNum, 32].Text;
                    var pointSymbTitle = !String.IsNullOrEmpty(pointSymbTitleString) ? Int32.Parse(pointSymbTitleString) : 0;

                    var titleYearString = worksheet.Cells[rowNum, 25].Text;
                    var titleYear = !String.IsNullOrEmpty(titleYearString) ? Int32.Parse(titleYearString.Substring(0, titleYearString.IndexOf('.'))) : 0;

                    var criticsScoreString = worksheet.Cells[rowNum, 4].Text;
                    var criticsScore = !String.IsNullOrEmpty(criticsScoreString) ? Double.Parse(criticsScoreString.Replace('.', ',')) : 0;

                    var durationString = worksheet.Cells[rowNum, 5].Text;
                    var duration = !String.IsNullOrEmpty(durationString) ? Double.Parse(durationString.Replace('.', ',')) : 0;

                    var directorFacebookLikesString = worksheet.Cells[rowNum, 6].Text;
                    var directorFacebookLikes = !String.IsNullOrEmpty(directorFacebookLikesString) ? Double.Parse(directorFacebookLikesString.Replace('.', ',')) : 0;

                    var actor3FacebookLikesString = worksheet.Cells[rowNum, 7].Text;
                    var actor3FacebookLikes = !String.IsNullOrEmpty(actor3FacebookLikesString) ? Double.Parse(actor3FacebookLikesString.Replace('.', ',')) : 0;

                    var actor1FacebookLikesString = worksheet.Cells[rowNum, 9].Text;
                    var actor1FacebookLikes = !String.IsNullOrEmpty(actor1FacebookLikesString) ? Double.Parse(actor1FacebookLikesString.Replace('.', ',')) : 0;

                    var grossString = worksheet.Cells[rowNum, 10].Text;
                    var gross = !String.IsNullOrEmpty(grossString) ? Double.Parse(grossString.Replace('.', ',')) : 0;

                    var budgetString = worksheet.Cells[rowNum, 24].Text;
                    var budget = !String.IsNullOrEmpty(budgetString) ? Double.Parse(budgetString.Replace('.', ',')) : 0;

                    var actor2FacebookLikesString = worksheet.Cells[rowNum, 26].Text;
                    var actor2FacebookLikes = !String.IsNullOrEmpty(actor2FacebookLikesString) ? Double.Parse(actor2FacebookLikesString.Replace('.', ',')) : 0;

                    var imdbScoreString = worksheet.Cells[rowNum, 27].Text;
                    var imdbScore = !String.IsNullOrEmpty(imdbScoreString) ? Double.Parse(imdbScoreString.Replace('.', ',')) : 0;

                    var ratioAspectString = worksheet.Cells[rowNum, 28].Text;
                    var ratioAspect = !String.IsNullOrEmpty(ratioAspectString) ? Double.Parse(ratioAspectString.Replace('.', ',')) : 0;

                    var unnamedString = worksheet.Cells[rowNum, 30].Text;
                    var unnamed = !String.IsNullOrEmpty(unnamedString) ? Double.Parse(unnamedString.Replace('.', ',')) : 0;

                    var color = worksheet.Cells[rowNum, 2].Text ?? String.Empty;
                    var directiorName = worksheet.Cells[rowNum, 3].Text ?? String.Empty;
                    var actor2Name = worksheet.Cells[rowNum, 8].Text ?? String.Empty;
                    var genre = worksheet.Cells[rowNum, 11].Text ?? String.Empty;
                    var actor1Name = worksheet.Cells[rowNum, 12].Text ?? String.Empty;
                    var movieTitle = worksheet.Cells[rowNum, 13].Text ?? String.Empty;
                    var actor3Name = worksheet.Cells[rowNum, 16].Text ?? String.Empty;
                    var plotKeywords = worksheet.Cells[rowNum, 18].Text ?? String.Empty;
                    var movieImdbLink = worksheet.Cells[rowNum, 19].Text ?? String.Empty;
                    var language = worksheet.Cells[rowNum, 21].Text ?? String.Empty;
                    var country = worksheet.Cells[rowNum, 22].Text ?? String.Empty;
                    var contentRating = worksheet.Cells[rowNum, 23].Text ?? String.Empty;

                    var movie = new Movie
                    {
                        Id = id,
                        Color = color,
                        DirectiorName = directiorName,
                        CriticsScore = criticsScore,
                        Duration = duration,
                        DirectorFacebookLikes = directorFacebookLikes,
                        Actor3FacebookLikes = actor3FacebookLikes,
                        Actor2Name = actor2Name,
                        Actor1FacebookLikes = actor1FacebookLikes,
                        Gross = gross,
                        Genre = new List<string>(genre.Split('|').Select(m => m.ToUpper())),
                        Actor1Name = actor1Name,
                        MovieTitle = movieTitle,
                        UsersVotesCount = usersVotesCount,
                        CastTotalFacebookLikes = castTotalFacebookLikes,
                        Actor3Name = actor3Name,
                        FaceNumInPoster = faceNumInPoster,
                        PlotKeywords = new List<string>(plotKeywords.Split('|').Select(m => m.ToUpper())),
                        MovieImdbLink = movieImdbLink,
                        UserViewsCount = userViewsCount,
                        Language = language,
                        Country = country,
                        ContentRating = contentRating,
                        Budget = budget,
                        TitleYear = titleYear,
                        Actor2FacebookLikes = actor2FacebookLikes,
                        ImdbScore = imdbScore,
                        RatioAspect = ratioAspect,
                        MovieFacebookLikes = movieFacebookLikes,
                        Unnamed = unnamed,
                        WarSymbTitle = warSymbTitle,
                        PointSymbTitle = pointSymbTitle
                    };

                    movies.Add(movie);
                }
            }

            return movies;
        }

        private static List<Movie> ValidateList(List<Movie> movies)
        {
            var incorrectActorNames = new List<string> { "Unit" };

            movies = movies
                .Where(m => incorrectActorNames.All(i => !m.Actor1Name.Contains(i)))
                .Where(m => incorrectActorNames.All(i => !m.Actor2Name.Contains(i)))
                .Where(m => incorrectActorNames.All(i => !m.Actor3Name.Contains(i)))
                .ToList();

            movies = movies
                .Where(m => !String.IsNullOrEmpty(m.MovieTitle))
                .ToList();

            return movies;
        }

        private static void GroupGenreByRating(List<Movie> movies)
        {
            var test = movies
                .SelectMany(m => m.Genre
                    .Select(i => new Tuple<string, double>(i.ToUpper(), m.ImdbScore)))
                .GroupBy(m => m.Item1)
                .OrderByDescending(group => group.Average(i => i.Item2));

            Console.WriteLine("Group by rating\n");

            foreach (var group in test)
            {

                Console.WriteLine($"{group.Key.PadRight(12)}: {String.Format("{0:n}", Math.Round(group.Average(m => m.Item2), 3))}");
            }
        }

        private static void GroupGenreByGross(List<Movie> movies)
        {
            var test = movies
                .SelectMany(m => m.Genre
                    .Select(i => new Tuple<string, double>(i.ToUpper(), m.Gross)))
                .GroupBy(m => m.Item1)
                .OrderByDescending(group => group.Average(i => i.Item2));

            Console.WriteLine("Group by gross\n");

            foreach (var group in test)
            {
                Console.WriteLine($"{group.Key.PadRight(12)}: {String.Format("{0:n}", Math.Round(group.Average(m => m.Item2), 3))}");
            }
        }

        private static void GroupGenreByLikes(List<Movie> movies)
        {
            var test = movies
                .SelectMany(m => m.Genre
                    .Select(i => new Tuple<string, double>(i.ToUpper(), m.CastTotalFacebookLikes)))
                .GroupBy(m => m.Item1)
                .OrderByDescending(group => group.Average(i => i.Item2));

            Console.WriteLine("Group by Facebook likes\n");

            foreach (var group in test)
            {
                Console.WriteLine($"{group.Key.PadRight(12)}: {String.Format("{0:n}", Math.Round(group.Average(m => m.Item2), 3))}");
            }
        }

        private static void GroupByActorDirectorIncome(List<Movie> movies)
        {
            var actor1List = movies
                .Select(a => new Tuple<string, string, double>(a.DirectiorName, a.Actor1Name, a.Gross));

            var actor2List = movies
                .Select(b => new Tuple<string, string, double>(b.DirectiorName, b.Actor2Name, b.Gross));

            var actor3List = movies
                .Select(c => new Tuple<string, string, double>(c.DirectiorName, c.Actor3Name, c.Gross));

            var atorDirectorIncome = actor1List.Concat(actor2List).Concat(actor3List).GroupBy(x => new { x.Item1, x.Item2 });

            Console.WriteLine("Group by director-actor gross\n");

            foreach (var group in atorDirectorIncome.OrderByDescending(m => m.Max(i => i.Item3)).Take(20))
            {
                Console.WriteLine($"Director: {group.Key.Item1}, Actor: {group.Key.Item2}, Gross: {String.Format("{0:n0}", group.Max(i => i.Item3))}");
            }
        }

        private static void GroupByCountryAverageRating(List<Movie> movies)
        {
            var scoreInCountry = movies
                .Where(m => !String.IsNullOrEmpty(m.Country))
                .GroupBy(m => m.Country);

            Console.WriteLine("Group by rating in country\n");

            foreach (var group in scoreInCountry.OrderByDescending(g => g.Average(i => i.ImdbScore)))
            {
                Console.WriteLine($"{group.Key.PadRight(15)}:\t{String.Format("{0:n}", Math.Round(group.Average(i => i.ImdbScore), 3))}");
            }
        }

        private static void GroupByDramaKeyWords(List<Movie> movies)
        {
            var dramaMovies = movies
                .Where(m => m.Genre.Contains("DRAMA"));

            var keywordsGroups = dramaMovies
                .SelectMany(m => m.PlotKeywords
                    .Where(pk => !String.IsNullOrEmpty(pk))
                    .Select(pk => new Tuple<string, int>(pk, m.PlotKeywords.IndexOf(pk))))
                .GroupBy(m => m.Item1);

            Console.WriteLine("Group by drama keywords\n");

            var averagePlotLength = dramaMovies.Average(m => m.PlotKeywords.Count);

            Console.WriteLine($"Average plot length {Math.Round(averagePlotLength, 2)}");
            Console.WriteLine();

            var plotMoveSB = new StringBuilder();

            plotMoveSB.Append("Plot movie: ");

            for (int i = 1; i < averagePlotLength; i++)
            {
                var plotKeyword = keywordsGroups
                    .Where(m => i == (int)m.Average(i => i.Item2))
                    .OrderByDescending(m => m.Count())
                    .FirstOrDefault();

                plotMoveSB.Append($"{plotKeyword.Key}=>");
            }

            Console.WriteLine(plotMoveSB.ToString().TrimEnd('>').TrimEnd('='));
            Console.WriteLine();

            //Console.WriteLine($"Average keywords more than {percent * 100}%:");
            //Console.WriteLine();

            Console.WriteLine("Keyword\t\t Count");

            foreach (var group in keywordsGroups.OrderByDescending(m => m.Count()))
            {
                Console.WriteLine($"{group.Key.PadRight(20)}:\t{group.Count()}");
                //Console.WriteLine($"{group.Key.PadRight(20)}:\t{group.Count()}\t\t{Math.Round(group.Average(m => m.Item2), 0)}");
            }

            //Console.WriteLine("**********************************");
        }

        private static void GroupByRating(List<Movie> movies)
        {
            var incorrectRatingNames = new List<string> { "Unrated", "Not Rated", "Approved" };

            var moviesRatingGroup = movies
                .Where(m => !String.IsNullOrEmpty(m.ContentRating) && m.Budget != 0 && !incorrectRatingNames.Contains(m.ContentRating))
                .GroupBy(m => m.ContentRating)
                .OrderByDescending(g => g.Count())
                .ToList();

            for (int i = 0; i < moviesRatingGroup.Count(); i++)
            {
                var group = moviesRatingGroup[i];

                var maxBudget = Math.Round(group.Max(m => m.Budget), 0);
                var minBudget = Math.Round(group.Min(m => m.Budget), 0);
                var averageBudget = Math.Round(group.Average(m => m.Budget), 0);

                var differenceWithPreviousString = "0%";

                if (i != 0)
                {
                    var previousGroup = moviesRatingGroup[i - 1];

                    var previousAverageBudget = previousGroup.Average(m => m.Budget);
                    var differenceWithPrevious = averageBudget - previousAverageBudget;

                    differenceWithPrevious = differenceWithPrevious / previousAverageBudget * 100;

                    differenceWithPreviousString = differenceWithPrevious > 0
                        ? $"+{Math.Round(differenceWithPrevious, 2)}%"
                        : $"{Math.Round(differenceWithPrevious, 2)}%";
                }

                Console.WriteLine($"{group.Key.PadRight(5)}" +
                    $"\t{group.Count()}" +
                    $"\t{String.Format("{0:n}", minBudget).PadRight(8)}" +
                    $"\t{String.Format("{0:n}", maxBudget).PadRight(8)}" +
                    $"\t{String.Format("{0:n}", averageBudget).PadRight(8)}" +
                    $"\t{differenceWithPreviousString.ToString().PadRight(8)}");
            }
        }

        private static void PredictNewFilm(List<Movie> movies)
        {
            var genre = "sci-fi";

            var percent = 0.005;

            //mostGrossSCIFIMovies
            var scifiMovies = movies.Where(m => m.Gross > 0 && m.Genre.Any(i => i.ToLower() == genre));

            #region Find Actor
            var actorMovies = movies.Where(m => m.Actor1FacebookLikes > 0 && m.Gross > 0);

            var actor = actorMovies
                .GroupBy(m => m.Actor1Name)
                .Where(g => g.Count() > (actorMovies.Count() * percent))
                .Select(g => new
                {
                    ActorName = g.Key,
                    MaxFacebookLikes = g.Max(m => m.Actor1FacebookLikes),
                    AverageFilmGross = g.Average(m => m.Gross),
                    FilmsCount = g.Count()
                })
                .OrderByDescending(m => m.MaxFacebookLikes)
                .FirstOrDefault();
            #endregion

            #region Find NotInGenreDirector
            var directorNotInGenreMovies = movies.Where(m => !String.IsNullOrEmpty(m.DirectiorName)
                && m.Gross > 0
                && m.Genre.All(i => i.ToLower() != genre));

            var notInGenreDirector = directorNotInGenreMovies
                .GroupBy(m => m.DirectiorName)
                .Where(g => g.Count() > (directorNotInGenreMovies.Count() * percent))
                .Select(g => new
                {
                    DirectorName = g.Key,
                    AverageImdbScore = g.Average(m => m.ImdbScore),
                    AverageFilmGross = g.Average(m => m.Gross),
                    FilmsCount = g.Count()
                })
                .OrderByDescending(m => m.AverageImdbScore)
                .FirstOrDefault();
            #endregion

            #region Find InGenreDirector
            var directorInGenreMovies = movies.Where(m => !String.IsNullOrEmpty(m.DirectiorName)
                && m.Gross > 0
                && m.Genre.Any(i => i.ToLower() == genre));

            var inGenreDirector = directorInGenreMovies
                .GroupBy(m => m.DirectiorName)
                .Where(g => g.Count() > (directorInGenreMovies.Count() * percent))
                .Select(g => new
                {
                    DirectorName = g.Key,
                    AverageImdbScore = g.Average(m => m.ImdbScore),
                    AverageFilmGross = g.Average(m => m.Gross),
                    FilmsCount = g.Count()
                })
                .OrderByDescending(m => m.AverageImdbScore)
                .FirstOrDefault();
            #endregion

            var totalSciFiAverageGross = scifiMovies.Average(m => m.Gross);
            Console.WriteLine($"Total SCI-FI films Count: {scifiMovies.Count()}\tAVG GROSS: {String.Format("{0:n0}", totalSciFiAverageGross)}\n");

            var differenceGrossForActor = (actor.AverageFilmGross - totalSciFiAverageGross) / totalSciFiAverageGross * 100;

            var differenceGrossForActorString = differenceGrossForActor > 0
                ? $"+{Math.Round(differenceGrossForActor, 2)}%"
                : $"{Math.Round(differenceGrossForActor, 2)}%";

            #region Predict for InGenreDirector
            {
                var differenceGrossForDirector = (inGenreDirector.AverageFilmGross - totalSciFiAverageGross) / totalSciFiAverageGross * 100;

                var differenceGrossForDirectorString = differenceGrossForDirector > 0
                    ? $"+{Math.Round(differenceGrossForDirector, 2)}%"
                    : $"{Math.Round(differenceGrossForDirector, 2)}%";

                Console.WriteLine("PREDICT FOR DIRECTOR WHO FILMED IN SCI-FI GENRE:\n");

                Console.WriteLine($"Director: {inGenreDirector.DirectorName.PadRight(20)}" +
                    $"\tFilms count: {inGenreDirector.FilmsCount}" +
                    $"\tAVG GROSS: {String.Format("{0:n0}", inGenreDirector.AverageFilmGross)}" +
                    $"\tGROSS CHANGE: {differenceGrossForDirectorString}");

                Console.WriteLine($"Actor: {actor.ActorName.PadRight(20)}" +
                    $"\tFilms count: {actor.FilmsCount}" +
                    $"\tAVG GROSS: {String.Format("{0:n0}", actor.AverageFilmGross)}" +
                    $"\tGROSS CHANGE: {differenceGrossForActorString}");

                var grossPredict = (totalSciFiAverageGross + actor.AverageFilmGross + inGenreDirector.AverageFilmGross) / 3;

                Console.WriteLine($"====GROSS PREDICT:\t{String.Format("{0:n0}", grossPredict)}====");
            }
            #endregion

            Console.WriteLine();
            Console.WriteLine();

            #region Predict for NotInGenreDirector
            {
                var differenceGrossForDirector = (notInGenreDirector.AverageFilmGross - totalSciFiAverageGross) / totalSciFiAverageGross * 100;

                var differenceGrossForDirectorString = differenceGrossForDirector > 0
                    ? $"+{Math.Round(differenceGrossForDirector, 2)}%"
                    : $"{Math.Round(differenceGrossForDirector, 2)}%";

                Console.WriteLine("PREDICT FOR DIRECTOR WHO DONT FILMED IN SCI-FI GENRE:\n");

                Console.WriteLine($"Director: {notInGenreDirector.DirectorName.PadRight(20)}" +
                    $"\tFilms count: {notInGenreDirector.FilmsCount}" +
                    $"\tAVG GROSS: {String.Format("{0:n0}", notInGenreDirector.AverageFilmGross)}" +
                    $"\tGROSS CHANGE: {differenceGrossForDirectorString}");

                Console.WriteLine($"Actor: {actor.ActorName.PadRight(20)}" +
                    $"\tFilms count: {actor.FilmsCount}" +
                    $"\tAVG GROSS: {String.Format("{0:n0}", actor.AverageFilmGross)}" +
                    $"\tGROSS CHANGE: {differenceGrossForActorString}");

                var grossPredict = (totalSciFiAverageGross + actor.AverageFilmGross + notInGenreDirector.AverageFilmGross) / 3;
                Console.WriteLine($"====GROSS PREDICT:\t{String.Format("{0:n0}", grossPredict)}====");
            }
            #endregion
        }
    }
}
