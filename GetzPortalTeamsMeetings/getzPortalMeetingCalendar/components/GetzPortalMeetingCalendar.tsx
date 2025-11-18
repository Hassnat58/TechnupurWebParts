import * as React from "react";
import styles from "./GetzPortalMeetingCalendar.module.scss";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

const GetzPortalMeetingCalendar = ({
  context,
}: {
  context: WebPartContext;
}) => {
  const [selectedDate, setSelectedDate] = React.useState(new Date());
  const [meetings, setMeetings] = React.useState<any>([]);
  const [loading, setloading] = React.useState(true);
  const getUserMeetings = async (): Promise<void> => {
    try {
      setloading(true);
      const startDate = new Date(selectedDate).toISOString();
      const endDate = new Date(
        new Date(selectedDate).setDate(new Date(selectedDate).getDate() + 1)
      ).toISOString();

      try {
        const client: MSGraphClientV3 =
          await context.msGraphClientFactory.getClient("3");
        console.log("Graph client retrieved successfully");

        client
          .api(
            `/me/calendar/calendarView?startDateTime=${startDate}&endDateTime=${endDate}`
          )
          .select("*,id,subject,start,end,location,attendees")
          .orderby("start/dateTime")
          .get((err: any, res: any) => {
            if (err) {
            setloading(false);
              console.error(err?.error, "meeting calendar");
              return;
            }

            setMeetings(res.value);
            setloading(false);
          });
      } catch (clientError) {
        console.error(
          clientError,
          "Error retrieving MSGraphClient - Permission Issue?"
        );
        setMeetings([]);
        setloading(false);
      }
    } catch (err) {
      setMeetings([]);
      setloading(false);
      console.error(err, "meeting calendar webpart");
    }
  };

  // use effect hook to fetch meetings each time meeting state is updated or the page reloads
  React.useEffect(() => {
    getUserMeetings().catch((err) =>{ 
      setloading(false); 
      console.error(err, "hi")} );
   

  }, [selectedDate]);

  const handlePrevDay = () => {
    setSelectedDate((prev) => new Date(prev.setDate(prev.getDate() - 1)));
  };

  const handleNextDay = () => {
    setSelectedDate((prev) => new Date(prev.setDate(prev.getDate() + 1)));
  };

  return (
    <div className={styles.meetingscalendar}>
      <div className={styles.header}>
        <div className={styles.titleDiv}>
          <svg
            width="20"
            height="22"
            viewBox="0 0 20 22"
            fill="none"
            xmlns="http://www.w3.org/2000/svg"
          >
            <path
              d="M2.75039 21.2917C2.20312 21.2917 1.73991 21.1021 1.36074 20.7229C0.981576 20.3438 0.791992 19.8806 0.791992 19.3333V4.83341C0.791992 4.28615 0.981576 3.82293 1.36074 3.44377C1.73991 3.0646 2.20312 2.87502 2.75039 2.87502H4.25026V0.583496H5.91697V2.87502H14.1254V0.583496H15.7504V2.87502H17.2503C17.7975 2.87502 18.2607 3.0646 18.6399 3.44377C19.0191 3.82293 19.2087 4.28615 19.2087 4.83341V19.3333C19.2087 19.8806 19.0191 20.3438 18.6399 20.7229C18.2607 21.1021 17.7975 21.2917 17.2503 21.2917H2.75039ZM2.75039 19.6667H17.2503C17.3337 19.6667 17.4101 19.6319 17.4794 19.5624C17.5489 19.4931 17.5837 19.4167 17.5837 19.3333V9.16675H2.41699V19.3333C2.41699 19.4167 2.45175 19.4931 2.52126 19.5624C2.5906 19.6319 2.66697 19.6667 2.75039 19.6667ZM2.41699 7.54175H17.5837V4.83341C17.5837 4.75 17.5489 4.67362 17.4794 4.60429C17.4101 4.53477 17.3337 4.50002 17.2503 4.50002H2.75039C2.66697 4.50002 2.5906 4.53477 2.52126 4.60429C2.45175 4.67362 2.41699 4.75 2.41699 4.83341V7.54175ZM10.0003 13.2501C9.73509 13.2501 9.50903 13.1567 9.32216 12.9698C9.13546 12.7831 9.04212 12.557 9.04212 12.2916C9.04212 12.0264 9.13546 11.8003 9.32216 11.6135C9.50903 11.4268 9.73509 11.3334 10.0003 11.3334C10.2656 11.3334 10.4916 11.4268 10.6785 11.6135C10.8652 11.8003 10.9585 12.0264 10.9585 12.2916C10.9585 12.557 10.8652 12.7831 10.6785 12.9698C10.4916 13.1567 10.2656 13.2501 10.0003 13.2501ZM5.66699 13.2501C5.40176 13.2501 5.1757 13.1567 4.98882 12.9698C4.80213 12.7831 4.70878 12.557 4.70878 12.2916C4.70878 12.0264 4.80213 11.8003 4.98882 11.6135C5.1757 11.4268 5.40176 11.3334 5.66699 11.3334C5.93223 11.3334 6.15828 11.4268 6.34516 11.6135C6.53185 11.8003 6.6252 12.0264 6.6252 12.2916C6.6252 12.557 6.53185 12.7831 6.34516 12.9698C6.15828 13.1567 5.93223 13.2501 5.66699 13.2501ZM14.3337 13.2501C14.0684 13.2501 13.8424 13.1567 13.6555 12.9698C13.4688 12.7831 13.3754 12.557 13.3754 12.2916C13.3754 12.0264 13.4688 11.8003 13.6555 11.6135C13.8424 11.4268 14.0684 11.3334 14.3337 11.3334C14.5989 11.3334 14.8249 11.4268 15.0118 11.6135C15.1985 11.8003 15.2919 12.0264 15.2919 12.2916C15.2919 12.557 15.1985 12.7831 15.0118 12.9698C14.8249 13.1567 14.5989 13.2501 14.3337 13.2501ZM10.0003 17.5C9.73509 17.5 9.50903 17.4066 9.32216 17.2197C9.13546 17.033 9.04212 16.807 9.04212 16.5418C9.04212 16.2764 9.13546 16.0503 9.32216 15.8636C9.50903 15.6768 9.73509 15.5833 10.0003 15.5833C10.2656 15.5833 10.4916 15.6768 10.6785 15.8636C10.8652 16.0503 10.9585 16.2764 10.9585 16.5418C10.9585 16.807 10.8652 17.033 10.6785 17.2197C10.4916 17.4066 10.2656 17.5 10.0003 17.5ZM5.66699 17.5C5.40176 17.5 5.1757 17.4066 4.98882 17.2197C4.80213 17.033 4.70878 16.807 4.70878 16.5418C4.70878 16.2764 4.80213 16.0503 4.98882 15.8636C5.1757 15.6768 5.40176 15.5833 5.66699 15.5833C5.93223 15.5833 6.15828 15.6768 6.34516 15.8636C6.53185 16.0503 6.6252 16.2764 6.6252 16.5418C6.6252 16.807 6.53185 17.033 6.34516 17.2197C6.15828 17.4066 5.93223 17.5 5.66699 17.5ZM14.3337 17.5C14.0684 17.5 13.8424 17.4066 13.6555 17.2197C13.4688 17.033 13.3754 16.807 13.3754 16.5418C13.3754 16.2764 13.4688 16.0503 13.6555 15.8636C13.8424 15.6768 14.0684 15.5833 14.3337 15.5833C14.5989 15.5833 14.8249 15.6768 15.0118 15.8636C15.1985 16.0503 15.2919 16.2764 15.2919 16.5418C15.2919 16.807 15.1985 17.033 15.0118 17.2197C14.8249 17.4066 14.5989 17.5 14.3337 17.5Z"
              fill="white"
            />
          </svg>

          <h3>&nbsp; Meetings Calendar</h3>
        </div>
      </div>
      <div className={styles.date}>
        <button className={styles.arrowbutton} onClick={handlePrevDay}>
          <svg
            width="8"
            height="12"
            viewBox="0 0 8 12"
            fill="none"
            xmlns="http://www.w3.org/2000/svg"
          >
            <path
              d="M3.40018 6L8.00018 1.4L6.60018 0L0.600183 6L6.60018 12L8.00018 10.6L3.40018 6Z"
              fill="#2A3440"
            />
          </svg>
        </button>
        <h2>{selectedDate.toDateString()}</h2>
        <button className={styles.arrowbutton} onClick={handleNextDay}>
          <svg
            width="8"
            height="12"
            viewBox="0 0 8 12"
            fill="none"
            xmlns="http://www.w3.org/2000/svg"
          >
            <path
              d="M4.59976 6L-0.000244141 1.4L1.39976 0L7.39976 6L1.39976 12L-0.000244141 10.6L4.59976 6Z"
              fill="#2A3440"
            />
          </svg>
        </button>
      </div>
      <div className={styles.meetingslist}>
        {!loading &&
         meetings.length > 0 ? (
          meetings.map((meeting: any, index: number) => (
            <div
              key={index}
              onClick={() => {
                if (meeting?.onlineMeeting?.joinUrl) {
                  window.open(meeting?.onlineMeeting?.joinUrl, "_blank");
                } else {
                  alert("This meeting doesn't have an online link.");
                }
              }}
              style={{ cursor: "pointer" }}
              title="Click to open in Teams"
              className={styles.meetingitem}
            >
              <div className={styles.time}>
                {" "}
                {
                  new Date(meeting.start.dateTime)
                    .toLocaleTimeString("en-US", {
                      hour: "numeric",
                      minute: "2-digit",
                    })
                    .split(" ")[0]
                }
                <br />
                {
                  new Date(meeting.start.dateTime)
                    .toLocaleTimeString("en-US", { hour12: true })
                    .split(" ")[1]
                }
              </div>
              <div className={styles.details}>
                <div className={styles.title}>{meeting.subject}</div>
                <div className={styles.description}>
                  {meeting.location.displayName || "Online"}
                </div>
              </div>
            </div>
          ))
        ) : loading ? (
          <div className={styles.nomeetings}>Loading...</div>
        ) : (
          <div className={styles.nomeetings}>
            No meetings scheduled for today.
          </div>
        )}
      </div>
    </div>
  );
};

export default GetzPortalMeetingCalendar;
