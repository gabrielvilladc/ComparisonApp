using AutoMapper;
using ComparisonApp.Entities;
using ComparisonApp.Entities.Performance;
using ComparisonApp.Entities.PortfolioStats;

namespace ComparisonApp.Mapping
{
    public  class SourceMappingProfile : Profile
    {
        public SourceMappingProfile() {

            //Holdings Report
            CreateMap<HoldingsAndPositionsReport, FinalHoldings>(MemberList.None)
                .ForMember(dest => dest.Source, opt => opt.MapFrom(src => src.ClosingPrice != null ? "CORICAPP" : "GetPositionsSP"))
                .ForMember(dest => dest.ClosingPrice, opt => opt.MapFrom(src => src.ClosingPrice != null ? src.ClosingPrice.Value.ToString("0.00") : src.Price != null ? src.Price.Value.ToString("0.00") : "0.00"))
                .ForMember(dest => dest.Collateral, opt => opt.MapFrom(src => src.Collateral != null ? src.Collateral : src.CollateralType != null ? src.CollateralType : string.Empty))
                .ForMember(dest => dest.Coupon, opt => opt.MapFrom(src => src.Coupon != null ? src.Coupon.Value.ToString("0.000") : "0.000"))
                .ForMember(dest => dest.CUSIP, opt => opt.MapFrom(src => src.Cusip != null && src.Cusip == "TOTAL" ? $"{src.Cusip} {char.ToUpper(src.CellLabel[0])}{src.CellLabel.Substring(1).ToLower()}" : string.IsNullOrEmpty(src.CUSIP) && src.Cusip == null ? $"TOTAL {src.SecurityType}" : src.Cusip ?? src.CUSIP))
                .ForMember(dest => dest.MarketValue, opt => opt.MapFrom(src => src.MarketValue != null ? src.MarketValue.Value.ToString("0.00") : "0.00"))
                .ForMember(dest => dest.MaturityDate, opt => opt.MapFrom(src => src.MaturityDate != null && src.MaturityDate != System.DateTime.MinValue ? src.MaturityDate.Value.ToString("MM/dd/yyyy") : src.Maturity != null && src.Maturity != System.DateTime.MinValue ? src.Maturity.Value.ToString("MM/dd/yyyy") : string.Empty))
                .ForMember(dest => dest.Moodys, opt => opt.MapFrom(src => src.Moodys ?? src.RatingMoodys ?? string.Empty))
                .ForMember(dest => dest.Paramount, opt => opt.MapFrom(src => src.ParAmount != null ? src.ParAmount.Value.ToString("0.00") : src.Paramount != null ? src.Paramount.Value.ToString("0.00") : "0.00"))
                .ForMember(dest => dest.SAndP, opt => opt.MapFrom(src => src.RatingSP ?? src.SAndP ?? string.Empty))
                .ForMember(dest => dest.SecurityType, opt => opt.MapFrom(src => src.CellLabel != null ? src.CellLabel : src.SecurityType))
                .ForMember(dest => dest.TsySpread, opt => opt.MapFrom(src => src.TsySpread != null ? src.TsySpread.Value.ToString("0.00") : "0.00"))
                .ForMember(dest => dest.YTW, opt => opt.MapFrom(src => src.YTW != null ? src.YTW.Value.ToString("0.00") : "0.00"));

            //Performance Mapping
            CreateMap<PerformanceAndCellData, FinalPerformance>(MemberList.None)
                .ForMember(dest => dest.Source, opt => opt.MapFrom(src => src.SinceInceptionReturn != null ? "CORICAPP" : "GetCellData"))
                .ForMember(dest => dest.SinceInceptionReturn, opt => opt.MapFrom(src => src.SinceInceptionReturn != null ? src.SinceInceptionReturn.Value.ToString("0.00") : src.SIDReturn != null ? src.SIDReturn.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.MTD, opt => opt.MapFrom(src => src.MTD != null ? src.MTD.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.QTD, opt => opt.MapFrom(src => src.QTD != null ? src.QTD.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.YTD, opt => opt.MapFrom(src => src.YTD != null ? src.YTD.Value.ToString("0.00") : string.Empty));

            //PortfolioStats Mapping
            //SecurityType = CellLabel
            //MarketValue = MarketValue
            //Weight = FeatureMarketWeight
            //ModifiedDuration = MDuration
            //CummReturn = SIDReturn
            CreateMap<PortfolioStatsAndCellData, FinalPortfolioStat>(MemberList.None)
                .ForMember(dest => dest.Source, opt => opt.MapFrom(src => src.SecurityType != null ? "CORICAPP" : "GetCellData"))
                .ForMember(dest => dest.SecurityType, opt => opt.MapFrom(src => src.CellLabel != null ? src.CellLabel : src.SecurityType))
                .ForMember(dest => dest.MarketValue, opt => opt.MapFrom(src => src.MarketValue != null ? src.MarketValue.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.Coupon, opt => opt.MapFrom(src => src.Coupon != null ? src.Coupon.Value.ToString("0.000") : string.Empty))
                .ForMember(dest => dest.BenchmarkCoupon, opt => opt.MapFrom(src => src.BenchmarkCoupon != null ? src.BenchmarkCoupon.Value.ToString("0.000") : string.Empty))
                .ForMember(dest => dest.Weight, opt => opt.MapFrom(src => src.FeatureMarketWeight != null ? src.FeatureMarketWeight.Value.ToString("0.00") : src.Weight != null ? src.Weight.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.BenchmarkWeight, opt => opt.MapFrom(src => src.BenchmarkWeight != null ? src.BenchmarkWeight.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.YTW, opt => opt.MapFrom(src => src.YTW != null ? src.YTW.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.BenchmarkYTW, opt => opt.MapFrom(src => src.BenchmarkYTW != null ? src.BenchmarkYTW.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.ModifiedDuration, opt => opt.MapFrom(src => src.MDuration != null ? src.MDuration.Value.ToString("0.00") : src.ModifiedDuration != null ? src.ModifiedDuration.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.BenchmarkModifiedDuration, opt => opt.MapFrom(src => src.BenchmarkModifiedDuration != null ? src.BenchmarkModifiedDuration.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.Convexity, opt => opt.MapFrom(src => src.Convexity != null ? src.Convexity.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.BenchmarkConvexity, opt => opt.MapFrom(src => src.BenchmarkConvexity != null ? src.BenchmarkConvexity.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.DailyReturn, opt => opt.MapFrom(src => src.DailyReturn != null ? src.DailyReturn.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.BenchmarkDailyReturn, opt => opt.MapFrom(src => src.BenchmarkDailyReturn != null ? src.BenchmarkDailyReturn.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.MTD, opt => opt.MapFrom(src => src.MTD != null ? src.MTD.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.BenchmarkMTD, opt => opt.MapFrom(src => src.BenchmarkMTD != null ? src.BenchmarkMTD.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.QTD, opt => opt.MapFrom(src => src.QTD != null ? src.QTD.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.BenchmarkQTD, opt => opt.MapFrom(src => src.BenchmarkQTD != null ? src.BenchmarkQTD.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.YTD, opt => opt.MapFrom(src => src.YTD != null ? src.YTD.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.BenchmarkYTD, opt => opt.MapFrom(src => src.BenchmarkYTD != null ? src.BenchmarkYTD.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.CummReturn, opt => opt.MapFrom(src => src.SIDReturn != null ? src.SIDReturn.Value.ToString("0.00") : src.CummReturn != null ? src.CummReturn.Value.ToString("0.00") : string.Empty))
                .ForMember(dest => dest.BenchmarkCummReturn, opt => opt.MapFrom(src => src.BenchmarkCummReturn != null ? src.BenchmarkCummReturn.Value.ToString("0.00") : string.Empty));
        }

            

    }
}
