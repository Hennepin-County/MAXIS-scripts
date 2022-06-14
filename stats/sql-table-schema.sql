CREATE TABLE [dbo].[usage_log](
       [ID] [bigint] IDENTITY(1,1) NOT NULL,
       [USERNAME] [varchar](255) NULL,
       [SDATE] [date] NULL,
       [STIME] [varchar](255) NULL,
       [SCRIPT_NAME] [varchar](255) NULL,
       [SRUNTIME] [decimal](18, 3) NULL,
       [CLOSING_MSGBOX] [varchar](max) NULL,
       [STATS_COUNTER] [int] NULL CONSTRAINT [DF_usage_log_STATS_COUNTER]  DEFAULT ((0)),
       [STATS_MANUALTIME] [int] NULL CONSTRAINT [DF_usage_log_STATS_MANUALTIME]  DEFAULT ((0)),
       [STATS_DENOMINATION] [varchar](255) NULL,
       [WORKER_COUNTY_CODE] [varchar](255) NULL,
       [SCRIPT_SUCCESS] [bit] NULL,
CONSTRAINT [PK_usage_log] PRIMARY KEY CLUSTERED 
(
       [ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [Data]
) ON [Data] TEXTIMAGE_ON [Data]

GO
