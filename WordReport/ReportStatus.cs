namespace WordReport
{
    using System;

    public enum ReportStatus
    {
        RS_START,
        RS_COMPLETEMAINDOC,
        RS_COMPLETEDETAIL,
        RS_READTOMERGEDATA,
        RS_WAITFORWORDTOSTART,
        RS_COMPLETEALL
    }
}
