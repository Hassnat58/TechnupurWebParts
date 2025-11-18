import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './Feelingtask.module.scss';
import { IFeelingtaskProps } from './IFeelingtaskProps';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

interface IEmojiOption {
  ResponseID: number;
  Response: string;
  staticUrl: string;
  animatedUrl: string;
  alt: string;
  videoUrl?: string;
}

interface IPollQuestion {
  id: number;
  question: string;
  questionsID: number;
  responses: string[];
  showEmoji: boolean;
}

const Feelingtask: React.FC<IFeelingtaskProps> = (props) => {
  const { context, spInstance } = props;

  const [questions, setQuestions] = useState<IPollQuestion[]>([]);
  const [selectedEmojis, setSelectedEmojis] = useState<{ [key: number]: number | null }>({});
  const [submittedQuestions, setSubmittedQuestions] = useState<{ [key: number]: boolean }>({});
  const [justSubmitted, setJustSubmitted] = useState<{ [key: number]: boolean }>({});
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [message, setMessage] = useState<{ text: string; type: MessageBarType } | null>(null);
  const [allSubmitted, setAllSubmitted] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isEmojiClicked, setIsEmojiClicked] = useState<{ [key: number]: boolean }>({});
  const [clickedEmoji, setClickedEmoji] = useState<{ [key: number]: IEmojiOption | null }>({});

  // Utility to add delay between API calls
  const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

  // Retry logic for handling throttling errors (HTTP 429)
  const fetchWithRetry = async (fetchFn: () => Promise<any>, retries = 3, delayMs = 5000): Promise<any> => {
    for (let i = 0; i < retries; i++) {
      try {
        return await fetchFn();
      } catch (error: any) {
        if (error.status === 429 && i < retries - 1) {
          const retryAfter = error.headers?.get("Retry-After") || delayMs / 1000;
          await delay(retryAfter * 1000);
          continue;
        }
        throw error;
      }
    }
  };

  useEffect(() => {
    const loadData = async () => {
      setIsLoading(true);
      await fetchListData();
    };
    loadData();
  }, []); // Empty dependency array to run only once on mount

  useEffect(() => {
    if (questions.length > 0) {
      const checkResponses = async () => {
        await checkTodayResponses();
        setIsLoading(false);
      };
      checkResponses();
    } else if (questions.length === 0 && !isLoading) {
      setIsLoading(false);
    }
  }, [questions]); // Run only when questions change

  const fetchListData = async (): Promise<void> => {
    try {
      const list = spInstance.web.lists.getByTitle("Poll");
      const items = await fetchWithRetry(() =>
        list.items
          .select("ID", "Questions", "QuestionsID", "Response", "ShowQuestions", "ShowEmoji")
          .filter("ShowQuestions eq 1")
          .orderBy("QuestionsID", true)()
      );

      await delay(1000); // Add a 1-second delay to prevent throttling

      if (items.length > 0) {
        const pollQuestions: IPollQuestion[] = items.map((item: any) => {
          let responses: string[] = [];
          if (Array.isArray(item.Response) && item.Response.length > 0) {
            responses = item.Response.map((r: string) => r.trim());
          } else {
            console.warn(`No valid Response for QuestionsID ${item.QuestionsID}:`, item.Response);
            responses = ["No responses available"];
          }

          return {
            id: item.ID,
            question: item.Questions,
            questionsID: item.QuestionsID,
            responses,
            showEmoji: item.ShowEmoji,
          };
        });

        setQuestions(pollQuestions);
      } else {
        console.error("No items found in Poll list with ShowQuestions set to Yes.");
        setQuestions([]);
      }
    } catch (error) {
      console.error("Error fetching list data:", error);
      setQuestions([]);
    }
  };

  const getEmojiOptions = (responses: string[]): IEmojiOption[] => {
    const emojiData = [
      { 
        staticUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f601/512.png", 
        animatedUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f601/512.gif", 
        alt: "Happy",
        videoUrl: "https://cldsl.sharepoint.com/sites/BusinessCentral/SiteAssets/smileVideo.mp4"
      },
      { 
        staticUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f613/512.png", 
        animatedUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f613/512.gif", 
        alt: "Sad",
        videoUrl: "https://cldsl.sharepoint.com/sites/BusinessCentral/SiteAssets/sadVideo.mp4"
      },
      { 
        staticUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f603/512.png", 
        animatedUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f603/512.gif", 
        alt: "Excited",
        videoUrl: "https://cldsl.sharepoint.com/sites/BusinessCentral/SiteAssets/excitedVideo.mp4"
      },
      { 
        staticUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f61e/512.png", 
        animatedUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f61e/512.gif", 
        alt: "Meh",
        videoUrl: "https://cldsl.sharepoint.com/sites/BusinessCentral/SiteAssets/mehVideo.mp4"
      },
      { 
        staticUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f621/512.png", 
        animatedUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f621/512.gif", 
        alt: "Frustrated",
        videoUrl: "https://cldsl.sharepoint.com/sites/BusinessCentral/SiteAssets/frustratedVideo.mp4"
      },
      { 
        staticUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f616/512.png", 
        animatedUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f616/512.gif", 
        alt: "Overwhelmed",
        videoUrl: "https://cldsl.sharepoint.com/sites/BusinessCentral/SiteAssets/overwhelmedVideo.mp4"
      },
      { 
        staticUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f615/512.png", 
        animatedUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f615/512.gif", 
        alt: "Confused",
        videoUrl: "https://cldsl.sharepoint.com/sites/BusinessCentral/SiteAssets/confusedVideo.mp4"
      },
      { 
        staticUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f634/512.png", 
        animatedUrl: "https://fonts.gstatic.com/s/e/notoemoji/latest/1f634/512.gif", 
        alt: "Tired",
        videoUrl: "https://cldsl.sharepoint.com/sites/BusinessCentral/SiteAssets/tiredVideo.mp4"
      },
    ];

    return responses.map((response: string, index: number) => ({
      ResponseID: index + 1,
      Response: response,
      staticUrl: emojiData[index % emojiData.length].staticUrl,
      animatedUrl: emojiData[index % emojiData.length].animatedUrl,
      alt: emojiData[index % emojiData.length].alt,
      videoUrl: emojiData[index % emojiData.length].videoUrl || undefined,
    }));
  };

  const checkTodayResponses = async (): Promise<void> => {
    try {
      const today = new Date();
      today.setHours(0, 0, 0, 0);

      const items = await fetchWithRetry(() =>
        spInstance.web.lists.getByTitle("PollResponse").items
          .filter(`ResponseDateTime ge datetime'${today.toISOString()}' and Status eq 'Submit'`)
          .select("ID", "UserResponse", "ResponseID", "User/Id", "User/Title", "User/EMail", "Status", "QuestionsId")
          .expand("User")()
      );

      await delay(1000); // Add a 1-second delay to prevent throttling

      const currentUserEmail: string = context.pageContext.user.email.toLowerCase();
      const userSubmissions = items.filter((item: any) =>
        item.User && item.User.EMail &&
        item.User.EMail.toLowerCase() === currentUserEmail
      );

      if (userSubmissions.length > 0) {
        const submitted: { [key: number]: boolean } = {};
        const selected: { [key: number]: number | null } = {};
        const clicked: { [key: number]: boolean } = {};
        const clickedEmojiData: { [key: number]: IEmojiOption | null } = {};

        // Map questions to their responses to get emoji options
        const questionResponseMap: { [key: number]: IEmojiOption[] } = {};
        questions.forEach((q) => {
          questionResponseMap[q.id] = getEmojiOptions(q.responses);
        });

        userSubmissions.forEach((submission: any) => {
          const questionId = submission.QuestionsId;
          submitted[questionId] = true;
          selected[questionId] = submission.ResponseID;

          // Set isEmojiClicked and clickedEmoji as if the emoji was clicked
          const emojiOptions = questionResponseMap[questionId];
          if (emojiOptions) {
            const selectedOption = emojiOptions.find((option: IEmojiOption) => option.ResponseID === submission.ResponseID);
            if (selectedOption) {
              clicked[questionId] = true;
              clickedEmojiData[questionId] = selectedOption;
            }
          }
        });

        setSubmittedQuestions(submitted);
        setSelectedEmojis(selected);
        setIsEmojiClicked(clicked);
        setClickedEmoji(clickedEmojiData);

        if (questions.length > 0 && questions.every(q => submitted[q.id])) {
          setAllSubmitted(true);
        }
      }
    } catch (error) {
      console.error("Error checking today's responses:", error);
    }
  };

  const handleEmojiSelect = (questionId: number, emoji: IEmojiOption): void => {
    if (!submittedQuestions[questionId]) {
      setSelectedEmojis(prev => ({
        ...prev,
        [questionId]: emoji.ResponseID,
      }));
      setIsEmojiClicked(prev => ({
        ...prev,
        [questionId]: true,
      }));
      setClickedEmoji(prev => ({
        ...prev,
        [questionId]: emoji,
      }));
    }
  };

  const handleBackClick = (questionId: number): void => {
    if (!submittedQuestions[questionId]) {
      setSelectedEmojis(prev => ({
        ...prev,
        [questionId]: null,
      }));
      setIsEmojiClicked(prev => ({
        ...prev,
        [questionId]: false,
      }));
      setClickedEmoji(prev => ({
        ...prev,
        [questionId]: null,
      }));
    }
  };

  const submitFeelingResponses = async (): Promise<void> => {
    const allQuestionsAnswered = questions.every(q => selectedEmojis[q.id] !== undefined && selectedEmojis[q.id] !== null);
    if (!allQuestionsAnswered) {
      setMessage({
        text: "Please Fill All Questions Response",
        type: MessageBarType.warning,
      });
      setTimeout(() => {
        setMessage(null);
      }, 3000);
      return;
    }

    setIsSubmitting(true);
    setMessage(null);

    try {
      const currentUser = await fetchWithRetry(() => spInstance.web.currentUser());
      const userId = currentUser.Id;

      for (const question of questions) {
        if (submittedQuestions[question.id]) continue;

        const selectedEmoji = selectedEmojis[question.id];
        const emojiOptions = getEmojiOptions(question.responses);
        const selectedOption: IEmojiOption | undefined = emojiOptions.find((option: IEmojiOption) => option.ResponseID === selectedEmoji);
        if (!selectedOption) throw new Error("Selected emoji not found");

        await fetchWithRetry(() =>
          spInstance.web.lists.getByTitle("PollResponse").items.add({
            UserResponse: selectedOption.Response,
            ResponseID: selectedEmoji,
            ResponseDateTime: new Date().toISOString(),
            Status: "Submit",
            UserId: userId,
            QuestionsId: question.id,
          })
        );

        await delay(1000); // Add a 1-second delay between submissions

        setSubmittedQuestions(prev => ({ ...prev, [question.id]: true }));
        setJustSubmitted(prev => ({ ...prev, [question.id]: true }));
      }

      // setMessage({
      //   text: "Thank you for your responses!",
      //   type: MessageBarType.success,
      // });

      setAllSubmitted(true);

      setTimeout(() => {
        setMessage(null);
        setJustSubmitted({});
      }, 3000);
    } catch (error) {
      console.error("Error submitting responses:", error);
      setMessage({
        text: "There was an error submitting your responses. Please try again.",
        type: MessageBarType.error,
      });

      setTimeout(() => {
        setMessage(null);
      }, 3000);
    } finally {
      setIsSubmitting(false);
    }
  };

  if (isLoading) {
    return <div>Loading...</div>;
  }

  return (
    <div className={styles.container}>
      {questions.length === 0 ? (
        <div className={styles.question}>No question available.</div>
      ) : (
        <>
          {questions.map((question) => {
            const selectedEmojiId = selectedEmojis[question.id];
            const emojiOptions = getEmojiOptions(question.responses);
            const selectedOption = emojiOptions.find((option: IEmojiOption) => option.ResponseID === selectedEmojiId);
            const displayText = submittedQuestions[question.id] && selectedOption
              ? `I am feeling ${selectedOption.Response} today!`
              : question.question;

            return (
              <div key={question.id} className={styles.questionBlock}>
                {isEmojiClicked[question.id] && clickedEmoji[question.id]?.videoUrl && (
                  <video className={styles.backgroundVideo} autoPlay loop muted>
                    <source src={clickedEmoji[question.id]?.videoUrl} type="video/mp4" />
                    Your browser does not support the video tag.
                  </video>
                )}
                {!submittedQuestions[question.id] && !allSubmitted && isEmojiClicked[question.id] && (
                  <button
                    className={styles.backButton}
                    onClick={() => handleBackClick(question.id)}
                    disabled={submittedQuestions[question.id] || allSubmitted}
                    aria-label="Go back to select another emoji"
                  >
                    ←
                  </button>
                )}
                <div className={styles.question}>{displayText}</div>

                {!submittedQuestions[question.id] && !allSubmitted && (
                  <div className={styles.emojiContainer}>
                    {getEmojiOptions(question.responses).map((option: IEmojiOption) => (
                      <div
                        key={option.ResponseID}
                        className={`${styles.emojiOption} ${selectedEmojis[question.id] === option.ResponseID ? styles.selected : ''} ${submittedQuestions[question.id] && selectedEmojis[question.id] !== option.ResponseID ? styles.disabled : ''} ${isEmojiClicked[question.id] && clickedEmoji[question.id]?.ResponseID !== option.ResponseID ? styles.hidden : ''}`}
                        onClick={() => handleEmojiSelect(question.id, option)}
                        aria-label={option.Response}
                        role="button"
                        tabIndex={0}
                      >
                        {question.showEmoji && (
                          <div className={styles.emoji}>
                            <picture>
                              <source srcSet={option.animatedUrl} type="image/gif" className={styles.animatedEmoji} />
                              <img src={option.staticUrl} alt={option.alt} width="36" height="36" className={styles.staticEmoji} />
                            </picture>
                          </div>
                        )}
                        <div className={styles.emojiName}>{option.Response}</div>
                      </div>
                    ))}
                  </div>
                )}

                {justSubmitted[question.id] && (
                  <div className={styles.submittedMessage}>
                    Thank you for submitting your response
                  </div>
                )}

                {submittedQuestions[question.id] && !justSubmitted[question.id] && (
                  <div className={styles.submittedMessage}>
                    You've already shared your response for this question. Thanks!
                  </div>
                )}
              </div>
            );
          })}

          {message && (
            <MessageBar
              messageBarType={message.type}
              isMultiline={false}
              dismissButtonAriaLabel="Close"
              className={styles.messageBar}
            >
              {message.text}
            </MessageBar>
          )}

          <div style={{ display: 'flex', gap: '10px', justifyContent: 'center'}}>
            {!allSubmitted && (
              <PrimaryButton
                text="Submit"
                onClick={submitFeelingResponses}
                disabled={isSubmitting}
                className={styles.submitButton}
              />
            )}
          </div>
        </>
      )}
    </div>
  );
};

export default Feelingtask;