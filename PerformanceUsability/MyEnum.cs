using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PerformanceUsability
{
    

    public enum ControlType
    {
        READY=0,
        SKIP, 
        PLAY,
        STOP,
        SETTING,
        FULLSCREEN,
        VOLUME
    }

    public enum DocType
    {
        DOC_READY = 100,
        DOC_TABLE,
        DOC_PICTURE,
        DOC_SHAPE,
        DOC_CHART,
        DOC_TYPING
    }

    public enum WebType
    {
        WEB_Chrome=0,
        WEB_FireFox,
        WEB_IE,
        WEB_EDGE
    }

    public enum MediaPlayType
    {
        MEDIA_WMP=0,
        MEDIA_MOVIE_AND_TV
    }

    public enum TaskStatus
    {
        TASK_RUNNING=0,
        TASK_FINISH
    }

    public enum TaskRunningList
    {
        TASK_IDLE=1,
        TASK_WEBACTOR,
        TASK_YOUTUBE,
        TASK_MEDIAPLAYER,
        TASK_STORAGE_ACTOR,
        TASK_DOCUMENT
    }

   

}
 