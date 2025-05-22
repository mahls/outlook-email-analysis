import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from wordcloud import WordCloud
from textblob import TextBlob
import stanza
import re
from collections import Counter
import os
from sklearn.feature_extraction.text import CountVectorizer
from wordcloud import STOPWORDS

# --- User-configurable variables ---
# These variables allow users to customize the dashboard's behavior and appearance.
DEFAULT_FALLBACK_FILENAME = "exported_emails.csv" # Default CSV file name if no file is uploaded
DEFAULT_DASHBOARD_TITLE = "Email Analysis Dashboard" # Default title for the Streamlit app
# Keywords to remove from email bodies (e.g., common signatures, company names)
# Users can add or remove terms here. Case-insensitive.
BODY_CLEANUP_KEYWORDS = ["kind regards", "best regards", "thanks", "thank you", "sincerely"]


# Download stanza English model (run once, will skip if already downloaded)
# This is placed outside functions to ensure it runs only once when the app starts.
try:
    stanza.download('en')
    # Initialize stanza pipeline for NER only
    nlp = stanza.Pipeline('en', processors='tokenize,ner', use_gpu=False)
except Exception as e:
    st.error(f"Failed to download Stanza model or initialize pipeline: {e}")
    st.info("Please ensure you have an active internet connection and try restarting the app.")
    st.stop()


st.set_page_config(page_title=DEFAULT_DASHBOARD_TITLE, layout="wide")

# Allow user to set the dashboard title
custom_dashboard_title = st.sidebar.text_input("Dashboard Title", DEFAULT_DASHBOARD_TITLE)
st.title(f"📨 {custom_dashboard_title}")


# File uploader and fallback logic
uploaded_file = st.file_uploader("Upload your exported email CSV/TSV file", type=["csv", "tsv"])
file_label = ""
data_source = None

if uploaded_file is not None:
    data_source = uploaded_file
    file_label = "📤 Uploaded file"
else:
    # Use the default fallback filename
    if os.path.exists(DEFAULT_FALLBACK_FILENAME):
        data_source = DEFAULT_FALLBACK_FILENAME
        file_label = f"📂 Loaded from {DEFAULT_FALLBACK_FILENAME}"
    else:
        st.error(f"No file uploaded and '{DEFAULT_FALLBACK_FILENAME}' not found in the current directory.")
        st.stop()

st.success(file_label)

# Load data
try:
    df = pd.read_csv(
        data_source,
        sep=None, # Automatically detect delimiter
        engine="python", # Use Python engine for flexible delimiter detection and on_bad_lines
        encoding="ISO-8859-1", # Common encoding for CSV exports
        on_bad_lines="skip" # Skip bad lines instead of raising an error
    )
except Exception as e:
    st.error(f"Error loading data: {e}")
    st.info("Please ensure your CSV/TSV file is correctly formatted and has the expected columns.")
    st.stop()


# Check for required columns
REQUIRED_COLUMNS = {'Subject', 'Sender', 'Date', 'Body'}
if REQUIRED_COLUMNS.issubset(df.columns):
    # Drop rows where any of the required columns are missing
    df.dropna(subset=list(REQUIRED_COLUMNS), inplace=True)

    # Convert 'Date' column to datetime objects
    # The date format is expected as "DD/MM/YYYY HH:MM:SS AM/PM"
    df['Date'] = pd.to_datetime(df['Date'], format="%d/%m/%Y %I:%M:%S %p", errors='coerce')
    df.dropna(subset=['Date'], inplace=True) # Remove rows where date parsing failed

    # Derive additional time-based features
    df['DateOnly'] = df['Date'].dt.date
    df['Hour'] = df['Date'].dt.hour
    df['Weekday'] = df['Date'].dt.day_name()
    df['Month'] = df['Date'].dt.to_period('M').astype(str)

    # Remove common footers and polite endings using the configurable keywords
    def clean_body(text, keywords_to_remove):
        text = str(text)
        # Create a regex pattern from the keywords_to_remove list
        # (?i) makes it case-insensitive
        # .* allows for any characters after the keyword (e.g., "neogen corp" if "neogen" is a keyword)
        pattern = r"(?i)(" + "|".join(re.escape(k) + ".*" for k in keywords_to_remove) + r")"
        text = re.sub(pattern, "", text)
        return text.strip() # Remove leading/trailing whitespace after cleaning

    # Apply body cleaning with user-defined keywords
    df['Body'] = df['Body'].apply(lambda x: clean_body(x, BODY_CLEANUP_KEYWORDS))

    # --- Sidebar Filters ---
    with st.sidebar:
        st.header("🔎 Filters")
        
        # Date range filter
        min_date = df['DateOnly'].min()
        max_date = df['DateOnly'].max()
        start_date = st.date_input("Start date", min_date)
        end_date = st.date_input("End date", max_date)

        # Sender filter (multiselect)
        all_senders = df['Sender'].unique().tolist()
        sender_filter = st.multiselect("Filter by sender", options=sorted(all_senders))

        # Keyword search in Subject or Body
        keyword = st.text_input("Search keywords in Subject or Body")
        
        # Allow user to add custom stop words for word cloud/common words analysis
        custom_stopwords_input = st.text_area("Add custom stop words (comma-separated)", "")
        CUSTOM_STOPWORDS = set(word.strip().lower() for word in custom_stopwords_input.split(',') if word.strip())
        ALL_STOPWORDS = STOPWORDS.union(CUSTOM_STOPWORDS)


    # Apply filters to the DataFrame
    filtered_df = df[(df['DateOnly'] >= start_date) & (df['DateOnly'] <= end_date)]
    if sender_filter:
        filtered_df = filtered_df[filtered_df['Sender'].isin(sender_filter)]
    if keyword:
        # Case-insensitive search across Subject and Body
        filtered_df = filtered_df[
            filtered_df['Subject'].str.contains(keyword, case=False, na=False) |
            filtered_df['Body'].str.contains(keyword, case=False, na=False)
        ]

    # --- Dashboard Content ---
    st.subheader("📊 Overview")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Emails", len(filtered_df))
    col2.metric("Unique Senders", filtered_df['Sender'].nunique())
    if not filtered_df.empty:
        col3.metric("Date Range", f"{filtered_df['DateOnly'].min()} → {filtered_df['DateOnly'].max()}")
    else:
        col3.metric("Date Range", "N/A")

    st.markdown("---")

    with st.expander("📈 Email Frequency by Date"):
        if not filtered_df.empty:
            fig = px.histogram(filtered_df, x="DateOnly", nbins=30, title="Email Volume Over Time")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.write("No data for selected filters.")

    with st.expander("📅 Heatmap: Emails by Hour & Day"):
        if not filtered_df.empty:
            # Group by Weekday and Hour, count, then unstack to create a pivot table
            heatmap_data = filtered_df.groupby(['Weekday', 'Hour']).size().unstack(fill_value=0)
            # Reindex to ensure consistent weekday order (Monday to Sunday)
            heatmap_data = heatmap_data.reindex(['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'])
            
            fig, ax = plt.subplots(figsize=(12, 5))
            sns.heatmap(heatmap_data, cmap="YlGnBu", ax=ax, annot=True, fmt="d", linewidths=.5)
            ax.set_title("Email Activity Heatmap (Emails by Hour & Day)", fontsize=14)
            ax.set_xlabel("Hour of Day", fontsize=12)
            ax.set_ylabel("Day of Week", fontsize=12)
            st.pyplot(fig)
        else:
            st.write("No data for selected filters.")

    with st.expander("📊 Monthly Trends"):
        if not filtered_df.empty:
            monthly_counts = filtered_df.groupby('Month').size().reset_index(name='Count')
            fig = px.line(monthly_counts, x='Month', y='Count', title='Monthly Email Volume')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.write("No data for selected filters.")

    with st.expander("👥 Top 20 Senders"):
        if not filtered_df.empty:
            top_senders = filtered_df['Sender'].value_counts().head(20)
            fig = px.bar(top_senders, x=top_senders.values, y=top_senders.index,
                         orientation='h', labels={'x':'Email Count', 'index':'Sender'},
                         title="Top 20 Email Senders")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.write("No data for selected filters.")

    with st.expander("🕵️ Outlier Detection (Email Bursts)"):
        if not filtered_df.empty:
            daily_counts = filtered_df.groupby('DateOnly').size()
            mean_count = daily_counts.mean()
            std_count = daily_counts.std()
            # Define a threshold for outliers (e.g., 2 standard deviations above the mean)
            threshold = mean_count + 2 * std_count
            burst_days = daily_counts[daily_counts > threshold]
            
            if not burst_days.empty:
                st.write("Detected burst days (unusually high email volume):")
                st.dataframe(burst_days.reset_index(name='Email Count').sort_values('Email Count', ascending=False))
            else:
                st.write("No significant burst days detected based on the 2-standard-deviation threshold.")
        else:
            st.write("No data for selected filters.")

    with st.expander("💬 Common Words in Subject Lines"):
        if not filtered_df.empty:
            def tokenize(text):
                # Clean text: remove non-alphanumeric, convert to lowercase
                text = re.sub(r"[^a-zA-Z0-9\s]", "", text.lower())
                # Split into words, filter by length and custom/default stopwords
                return [word for word in text.split() if len(word) > 2 and word not in ALL_STOPWORDS]

            words = Counter()
            # Apply tokenization to each subject line and update the counter
            filtered_df['Subject'].dropna().apply(lambda s: words.update(tokenize(s)))
            
            if words:
                common_words = pd.DataFrame(words.most_common(20), columns=['Word', 'Count'])
                fig = px.bar(common_words, x='Count', y='Word', orientation='h', title="Top 20 Common Words in Subject Lines")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.write("No common words found after filtering.")
        else:
            st.write("No data for selected filters.")

    with st.expander("📚 Top Bigrams in Subject Lines"):
        if not filtered_df.empty:
            # Use CountVectorizer to find bigrams (sequences of 2 words)
            vectorizer = CountVectorizer(ngram_range=(2, 2), stop_words=list(ALL_STOPWORDS))
            # Fit and transform the subject lines
            X = vectorizer.fit_transform(filtered_df['Subject'].fillna(""))
            
            # Sum the occurrences of each bigram
            sum_words = X.sum(axis=0)
            # Create a list of (bigram, count) tuples
            words_freq = [(word, sum_words[0, idx]) for word, idx in vectorizer.vocabulary_.items()]
            # Sort by count in descending order and get the top 20
            sorted_bigrams = sorted(words_freq, key=lambda x: x[1], reverse=True)[:20]
            
            if sorted_bigrams:
                bigram_df = pd.DataFrame(sorted_bigrams, columns=['Bigram', 'Count'])
                fig = px.bar(bigram_df, x='Count', y='Bigram', orientation='h', title="Top 20 Subject Line Bigrams")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.write("No significant bigrams found after filtering.")
        else:
            st.write("No data for selected filters.")

    with st.expander("☁️ Word Cloud of Email Bodies"):
        if not filtered_df.empty:
            # Concatenate all cleaned email bodies into a single string
            all_text = " ".join(filtered_df['Body'].astype(str))
            # Generate word cloud
            wc = WordCloud(width=1000, height=400, background_color='white', stopwords=ALL_STOPWORDS).generate(all_text)
            
            fig, ax = plt.subplots(figsize=(12, 5))
            ax.imshow(wc, interpolation='bilinear')
            ax.axis("off") # Hide axes
            ax.set_title("Word Cloud of Email Bodies", fontsize=14)
            st.pyplot(fig)
        else:
            st.write("No data for selected filters.")

    with st.expander("🔍 Named Entity Recognition (NER)"):
        if not filtered_df.empty:
            # Take a sample of bodies to avoid processing extremely large text, which can be slow
            # Join a sample of bodies for NER processing
            sample_text = " ".join(filtered_df['Body'].sample(min(50, len(filtered_df))).tolist())
            
            # Process text with Stanza pipeline
            doc = nlp(sample_text)

            ents = []
            # Extract entities from each sentence
            for sent in doc.sentences:
                for ent in sent.ents:
                    ents.append((ent.text, ent.type))
            
            if ents:
                entity_df = pd.DataFrame(ents, columns=['Entity', 'Label'])
                st.markdown("#### Top Named Entities Found (Sample of Email Bodies)")
                st.dataframe(entity_df.value_counts().reset_index(name='Count').head(20))
            else:
                st.write("No named entities found in the sample email bodies.")
        else:
            st.write("No data for selected filters.")

    with st.expander("📈 Sentiment Analysis"):
        if not filtered_df.empty:
            # Calculate sentiment polarity for each email body
            filtered_df['Polarity'] = filtered_df['Body'].apply(lambda x: TextBlob(str(x)).sentiment.polarity)
            
            fig = px.histogram(filtered_df, x='Polarity', nbins=20, title="Email Sentiment Distribution")
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("**📉 Top 10 Most Negative Emails**")
            # Display top 10 most negative emails, ordered by polarity
            st.dataframe(filtered_df.sort_values('Polarity').head(10)[['Date', 'Sender', 'Subject', 'Polarity']])

            st.markdown("**📈 Top 10 Most Positive Emails**")
            # Display top 10 most positive emails, ordered by polarity
            st.dataframe(filtered_df.sort_values('Polarity', ascending=False).head(10)[['Date', 'Sender', 'Subject', 'Polarity', 'Body']])
        else:
            st.write("No data for selected filters.")

    # -------- Dynamic Summary & Behavioral Insights --------
    with st.expander("✅ Dynamic Summary & Behavioral Insights"):
        if filtered_df.empty:
            st.markdown("No data available for the selected filters to generate insights.")
        else:
            # Polarity groups for insights
            positives = filtered_df[filtered_df['Polarity'] > 0.5]
            negatives = filtered_df[filtered_df['Polarity'] < -0.5]

            st.markdown(f"- **Total Emails Analyzed**: {len(filtered_df)}")
            st.markdown(f"- **Highly Positive Emails (Polarity > 0.5)**: {len(positives)}")
            st.markdown(f"- **Highly Negative Emails (Polarity < -0.5)**: {len(negatives)}")

            # Top positive senders
            if not positives.empty:
                top_positive_senders = positives['Sender'].value_counts().head(5)
                st.markdown("### 👍 Top Positive Senders")
                for sender, count in top_positive_senders.items():
                    st.markdown(f"- {sender}: {count} positive emails")
            else:
                st.markdown("### 👍 No highly positive emails detected.")

            # Top negative senders
            if not negatives.empty:
                top_negative_senders = negatives['Sender'].value_counts().head(5)
                st.markdown("### ⚠️ Top Negative Senders")
                for sender, count in top_negative_senders.items():
                    st.markdown(f"- {sender}: {count} negative emails")
            else:
                st.markdown("### ⚠️ No highly negative emails detected.")

            # Burst days detection
            daily_counts = filtered_df.groupby('DateOnly').size()
            mean_count = daily_counts.mean()
            std_count = daily_counts.std()
            threshold = mean_count + 2 * std_count # 2 standard deviations above mean for burst
            burst_days = daily_counts[daily_counts > threshold]

            if not burst_days.empty:
                st.markdown(f"### 🔥 Burst Days Detected: {len(burst_days)} day(s) with unusually high email volume")
                for date, count in burst_days.items():
                    st.markdown(f"- {date}: {count} emails")
            else:
                st.markdown("### 🔥 No unusual burst days detected.")

            # Behavioral and investigative notes

            # 1. Sender behavior patterns
            sender_email_counts = filtered_df['Sender'].value_counts()
            if not sender_email_counts.empty:
                top_sender = sender_email_counts.idxmax()
                top_sender_count = sender_email_counts.max()

                st.markdown("### 🔎 Sender Behavior Patterns")
                st.markdown(f"- The top sender is **{top_sender}** with **{top_sender_count}** emails.")
                if top_sender_count > 0.3 * len(filtered_df): # If top sender accounts for >30% of emails
                    st.markdown("  - This suggests a dominant communication source, possibly a key stakeholder or automated system.")
                else:
                    st.markdown("  - Email distribution is relatively balanced among senders.")
            else:
                st.markdown("### 🔎 No sender behavior patterns to analyze.")


            # 2. Temporal communication patterns
            if not filtered_df.empty:
                peak_hour = filtered_df['Hour'].mode()[0]
                peak_weekday = filtered_df['Weekday'].mode()[0]
                st.markdown("### 🕒 Temporal Communication Patterns")
                st.markdown(f"- Peak email sending hour is around **{peak_hour}:00**.")
                st.markdown(f"- Most active day of the week is **{peak_weekday}**.")
                st.markdown("- These patterns may reflect typical business hours and weekday workload.")
            else:
                st.markdown("### 🕒 No temporal communication patterns to analyze.")


            # 3. Sentiment dynamics
            if not filtered_df.empty:
                avg_polarity = filtered_df['Polarity'].mean()
                st.markdown("### 📊 Sentiment Dynamics")
                if avg_polarity > 0.1:
                    st.markdown("- Overall sentiment is slightly positive, indicating generally constructive communication.")
                elif avg_polarity < -0.1:
                    st.markdown("- Overall sentiment leans negative, possible signs of dissatisfaction or conflict.")
                else:
                    st.markdown("- Sentiment is mostly neutral, reflecting balanced communication.")
            else:
                st.markdown("### 📊 No sentiment dynamics to analyze.")


            # 4. Linguistic cues & investigation flags
            word_counts = Counter()
            # Use the configurable BODY_CLEANUP_KEYWORDS for common investigative terms as well
            # You might want a separate list for "investigative terms" if they are different from cleanup keywords.
            # For now, let's use a generic list of common terms that might flag issues.
            common_investigative_terms = ['issue', 'problem', 'urgent', 'delay', 'fail', 'error', 'complaint', 'request', 'bug', 'fix', 'escalate']

            # Update word counts from cleaned body text
            filtered_df['Body'].dropna().apply(lambda text: word_counts.update(re.findall(r'\b\w+\b', text.lower())))
            
            flagged_terms = {term: word_counts[term] for term in common_investigative_terms if word_counts[term] > 0}

            st.markdown("### 🧐 Investigative Linguistic Cues")
            if flagged_terms:
                for term, count in flagged_terms.items():
                    st.markdown(f"- Term **'{term}'** appeared **{count}** times, suggesting frequent mentions of potential issues or concerns.")
                st.markdown("- These keywords could indicate areas requiring operational review or support focus.")
            else:
                st.markdown("- No strong presence of common investigative or complaint-related keywords detected.")


            # 5. Polite or evasive language detection (basic heuristic)
            # Use the configurable BODY_CLEANUP_KEYWORDS for polite phrases as well, or define a separate list
            polite_phrases_for_detection = ["kind regards", "best regards", "thank you", "thanks", "sincerely", "please", "appreciate"]

            polite_count = sum(filtered_df['Body'].str.lower().str.contains('|'.join(re.escape(p) for p in polite_phrases_for_detection), na=False))
            polite_ratio = polite_count / len(filtered_df) if len(filtered_df) > 0 else 0

            st.markdown("### 🙏 Politeness & Communication Tone")
            st.markdown(f"- Approximately **{polite_ratio:.1%}** of emails include polite or formal phrases.")
            if polite_ratio > 0.4:
                st.markdown("- High politeness ratio may suggest formal business communications or attempts to soften requests.")
            else:
                st.markdown("- Lower politeness usage could indicate more direct or informal exchanges.")

            # 6. Anomaly notes
            # Refined anomaly detection logic
            if not burst_days.empty or polite_ratio < 0.2 or avg_polarity < -0.1:
                st.markdown("### 🚩 Potential Anomalies or Areas for Further Investigation")
                if not burst_days.empty:
                    st.markdown("- **Email Bursts**: Unusually high email volume detected on certain days. Investigate these periods for specific events or issues.")
                if polite_ratio < 0.2:
                    st.markdown("- **Low Politeness**: A low ratio of polite phrases might indicate a more direct, urgent, or potentially confrontational tone in communications.")
                if avg_polarity < -0.1:
                    st.markdown("- **Negative Sentiment**: Overall negative sentiment suggests potential dissatisfaction or ongoing problems that might need attention.")
            else:
                st.markdown("### ✅ Communication appears consistent with typical patterns. No significant anomalies detected.")

else:
    st.error("Your file is missing one or more required columns: Subject, Sender, Date, Body.")
    st.info("Please ensure your CSV/TSV file has these columns.")