import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from collections import defaultdict
from datetime import datetime, timedelta
import pickle
import argparse
from typing import Dict, List, Tuple, Optional
from rich.console import Console
from rich.table import Table
import pandas as pd
from plotly import graph_objects as go
from plotly.subplots import make_subplots
import time
import re


SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive.metadata',
    'https://www.googleapis.com/auth/drive.activity',
    'https://www.googleapis.com/auth/drive.activity.readonly'
]

class GoogleDocsAnalyzer:
    def __init__(self):
        self.creds: Optional[Credentials] = None
        self.service = None
        self.drive_service = None
        self.activity_service = None
        self.console = Console()

    def authenticate(self) -> None:
        """Handle Google OAuth authentication."""
        try:
            if os.path.exists('token.pickle'):
                with open('token.pickle', 'rb') as token:
                    self.creds = pickle.load(token)
            
            if not self.creds or not self.creds.valid:
                if self.creds and self.creds.expired and self.creds.refresh_token:
                    self.creds.refresh(Request())
                else:
                    if not os.path.exists('credentials.json'):
                        raise FileNotFoundError("credentials.json not found. Please download it from Google Cloud Console.")
                    
                    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                    self.creds = flow.run_local_server(port=0)
                
                with open('token.pickle', 'wb') as token:
                    pickle.dump(self.creds, token)

            # Build services
            self.service = build('docs', 'v1', credentials=self.creds)
            self.drive_service = build('drive', 'v3', credentials=self.creds)
            self.activity_service = build('driveactivity', 'v2', credentials=self.creds)
            
            # Verify authentication
            about = self.drive_service.about().get(fields="user").execute()
            self.console.print(f"[green]Successfully authenticated as: {about['user']['emailAddress']}[/green]")
            
        except Exception as e:
            self.console.print(f"[red]Authentication failed: {str(e)}[/red]")
            raise

    def get_document_metadata(self, doc_id: str) -> Dict[str, str]:
        """Retrieve basic document metadata."""
        try:
            document = self.service.documents().get(documentId=doc_id).execute()
            file_metadata = self.drive_service.files().get(
                fileId=doc_id, 
                fields="createdTime,modifiedTime,owners,lastModifyingUser"
            ).execute()
            
            return {
                'title': document.get('title', 'Untitled'),
                'created_time': self._format_timestamp(file_metadata.get('createdTime', 'Unknown')),
                'modified_time': self._format_timestamp(file_metadata.get('modifiedTime', 'Unknown')),
                'owner': file_metadata.get('owners', [{}])[0].get('displayName', 'Unknown'),
                'last_modifier': file_metadata.get('lastModifyingUser', {}).get('displayName', 'Unknown')
            }
        except Exception as e:
            self.console.print(f"[red]Error retrieving document metadata: {str(e)}[/red]")
            return {}

    def _format_timestamp(self, timestamp: str) -> str:
        """Convert ISO timestamp to readable format."""
        if timestamp == 'Unknown':
            return timestamp
        try:
            dt = datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except:
            return timestamp

    def get_revision_history(self, doc_id: str) -> List[dict]:
        """Get the complete revision history of the document."""
        try:
            self.console.print("[yellow]Fetching revision history...[/yellow]")
            revisions = []
            page_token = None
            
            while True:
                response = self.drive_service.revisions().list(
                    fileId=doc_id,
                    pageSize=100,
                    fields="nextPageToken,revisions(id,modifiedTime,lastModifyingUser)",
                    pageToken=page_token
                ).execute()
                
                revisions.extend(response.get('revisions', []))
                page_token = response.get('nextPageToken')
                if not page_token:
                    break
            
            self.console.print(f"[green]Successfully retrieved {len(revisions)} revisions[/green]")
            return revisions
        except Exception as e:
            self.console.print(f"[red]Error retrieving revision history: {str(e)}[/red]")
            return []

    def get_activity_history(self, doc_id: str, lookback_days: int = 365) -> List[dict]:
        """Get detailed activity history including content changes."""
        try:
            activities = []
            page_token = None
            start_time = (datetime.utcnow() - timedelta(days=lookback_days)).isoformat() + 'Z'
            
            while True:
                results = self.activity_service.activity().query(
                    body={
                        'itemName': f'items/{doc_id}',
                        'pageSize': 100,
                        'pageToken': page_token,
                        'filter': f'time >= "{start_time}"'
                    }
                ).execute()
                
                activities.extend(results.get('activities', []))
                page_token = results.get('nextPageToken')
                
                if not page_token:
                    break
            
            return activities
        except Exception as e:
            self.console.print(f"[red]Error retrieving activity history: {str(e)}[/red]")
            return []



    def get_comments(self, doc_id: str) -> List[dict]:
        """Retrieve all comments and their metadata."""
        try:
            self.console.print("[yellow]Fetching comments...[/yellow]")
            comments = []
            page_token = None
            
            while True:
                response = self.drive_service.comments().list(
                    fileId=doc_id,
                    fields="nextPageToken,comments(id,content,author,createdTime,resolved,modifiedTime,replies)",
                    includeDeleted=False,
                    pageToken=page_token
                ).execute()
                
                comments.extend(response.get('comments', []))
                page_token = response.get('nextPageToken')
                if not page_token:
                    break
                
            self.console.print(f"[green]Successfully retrieved {len(comments)} comments[/green]")
            return comments
        except Exception as e:
            self.console.print(f"[red]Error retrieving comments: {str(e)}[/red]")
            return []

    def _get_user_email(self, user_info: dict) -> str:
        """Extract email address from user info dictionary."""
        try:
            if 'emailAddress' in user_info:
                return user_info['emailAddress']
            return 'Email not available'
        except:
            return 'Email not available'

    # Update analyze_contributions method
    def analyze_contributions(self, revisions: List[dict]) -> Dict[str, dict]:
        """Analyze user contributions from revision history."""
        contributions = defaultdict(lambda: {
            'revision_count': 0,
            'last_modified': None,
            'first_modified': None,
            'email': 'Email not available'
        })

        for revision in sorted(revisions, key=lambda x: x.get('modifiedTime', '')):
            user_info = revision.get('lastModifyingUser', {})
            user = user_info.get('displayName', 'Unknown User')
            email = self._get_user_email(user_info)
            mod_time = self._format_timestamp(revision.get('modifiedTime', ''))
            
            contributions[user]['revision_count'] += 1
            contributions[user]['email'] = email
            
            if not contributions[user]['first_modified']:
                contributions[user]['first_modified'] = mod_time
            contributions[user]['last_modified'] = mod_time

        return dict(contributions)

    def calculate_word_contributions(self, activities: List[dict]) -> Dict[str, int]:
        """Calculate approximate word contributions per user based on activity history."""
        word_counts = defaultdict(int)
        
        try:
            for activity in activities:
                # Check if activity has targets and the first target has a driveItem
                if ('targets' in activity and 
                    activity['targets'] and 
                    'driveItem' in activity['targets'][0]):
                    
                    actions = activity.get('primaryActionDetail', {})
                    if 'edit' in actions:
                        actor = self._get_actor_name(activity)
                        
                        # Get edit details
                        edit_details = actions['edit']
                        
                        # If it's a suggestion, skip it
                        if edit_details.get('suggestion', False):
                            continue
                            
                        # Count different types of edits
                        if 'documentChange' in edit_details:
                            # Major document changes
                            word_counts[actor] += 10
                        elif 'delete' in edit_details:
                            # Deletion operations
                            word_counts[actor] += 3
                        else:
                            # Regular edits
                            word_counts[actor] += 5
                
        except Exception as e:
            self.console.print(f"[red]Error calculating word contributions: {str(e)}[/red]")
            # Print the problematic activity for debugging
            if 'activity' in locals():
                self.console.print(f"[yellow]Problematic activity: {activity}[/yellow]")
        
        return dict(word_counts)

    def get_revision_content(self, doc_id: str, revision_id: str) -> str:
        """Get the content of a specific revision, with proper delay and error handling."""
        try:
            import time
            import requests
            
            # Get revision metadata
            revision = self.drive_service.revisions().get(
                fileId=doc_id,
                revisionId=revision_id,
                fields="exportLinks,modifiedTime"
            ).execute()
            
            # Get the plain text export link
            export_links = revision.get('exportLinks', {})
            text_link = export_links.get('text/plain')
            
            if not text_link:
                return ""
            
            # Add delay to avoid overloading
            time.sleep(5)  # 3 second delay between requests
            
            # Get the content
            headers = {'Authorization': f'Bearer {self.creds.token}'}
            response = requests.get(text_link, headers=headers)
            
            if response.status_code == 200:
                return response.text
            else:
                self.console.print(f"[yellow]Warning: Failed to get content for revision {revision_id}: {response.status_code}[/yellow]")
                return ""
                
        except Exception as e:
            self.console.print(f"[red]Error retrieving revision content: {str(e)}[/red]")
            return ""

    def _get_actor_name(self, activity: dict) -> str:
        """Extract actor name from activity."""
        try:
            actor = activity.get('actors', [{}])[0]
            if 'user' in actor:
                return actor['user'].get('knownUser', {}).get('personName', 'Unknown User')
            return 'Unknown User'
        except:
            return 'Unknown User'

    def create_historical_analysis(self, activities: List[dict], comments: List[dict]) -> pd.DataFrame:
        """Create time series data for various metrics."""
        metrics = defaultdict(lambda: defaultdict(int))
        
        # Process activities with daily and weekly aggregation
        for activity in activities:
            timestamp = activity.get('timestamp', '')
            if timestamp:
                date = datetime.fromisoformat(timestamp.replace('Z', '+00:00')).date()
                metrics[date]['edits'] += 1
                
        # Process comments and their lifecycle
        for comment in comments:
            created_time = comment.get('createdTime', '')
            if created_time:
                date = datetime.fromisoformat(created_time.replace('Z', '+00:00')).date()
                metrics[date]['comments'] += 1
                
                # Track replies
                for reply in comment.get('replies', []):
                    reply_time = reply.get('createdTime', '')
                    if reply_time:
                        reply_date = datetime.fromisoformat(reply_time.replace('Z', '+00:00')).date()
                        metrics[reply_date]['replies'] += 1
                
                # Track resolutions
                if comment.get('resolved', False):
                    resolved_time = comment.get('modifiedTime', created_time)
                    resolved_date = datetime.fromisoformat(resolved_time.replace('Z', '+00:00')).date()
                    metrics[resolved_date]['resolved'] += 1
        
        # Create DataFrame with complete date range
        if metrics:
            date_range = pd.date_range(min(metrics.keys()), max(metrics.keys()))
            data = []
            for date in date_range:
                date_metrics = metrics[date.date()]
                data.append({
                    'date': date.date(),
                    'edits': date_metrics['edits'],
                    'comments': date_metrics['comments'],
                    'replies': date_metrics['replies'],
                    'resolved': date_metrics['resolved']
                })
            return pd.DataFrame(data)
        return pd.DataFrame()


    def plot_historical_metrics(self, df: pd.DataFrame, title: str, word_growth_df: Optional[pd.DataFrame] = None) -> None:
        """Create and save an interactive plot of historical metrics.
        
        Args:
            df: DataFrame containing comments and edits history
            title: Title for the plot
            word_growth_df: Optional DataFrame containing word count history
        """
        if df.empty:
            self.console.print("[yellow]No historical data available for plotting[/yellow]")
            return

        # Determine subplot layout
        rows = 3 if word_growth_df is not None and not word_growth_df.empty else 2
        subplot_titles = ['Document Edits', 'Comments Made', 'Replies Made', 'Comments Resolved']
        if word_growth_df is not None and not word_growth_df.empty:
            subplot_titles.extend(['Word Count Growth', 'Word Changes'])
        
        fig = make_subplots(
            rows=rows, 
            cols=2,
            subplot_titles=subplot_titles,
            vertical_spacing=0.12
        )
        
        # Add activity traces
        traces = [
            ('edits', 'Edits', 'rgb(31, 119, 180)', 1, 1),
            ('comments', 'Comments', 'rgb(255, 127, 14)', 1, 2),
            ('replies', 'Replies', 'rgb(44, 160, 44)', 2, 1),
            ('resolved', 'Resolved', 'rgb(214, 39, 40)', 2, 2)
        ]
        
        for metric, name, color, row, col in traces:
            fig.add_trace(
                go.Scatter(
                    x=df['date'],
                    y=df[metric],
                    mode='lines+markers',
                    name=name,
                    line=dict(width=2, color=color)
                ),
                row=row, col=col
            )
        
        # Add word growth charts if available
        if word_growth_df is not None and not word_growth_df.empty:
            # Total word count over time
            fig.add_trace(
                go.Scatter(
                    x=word_growth_df['timestamp'],
                    y=word_growth_df['total_words'],
                    mode='lines+markers',
                    name='Total Words',
                    line=dict(width=2, color='rgb(148, 103, 189)')
                ),
                row=3, col=1
            )
            
            # Word changes per revision
            fig.add_trace(
                go.Bar(
                    x=word_growth_df['timestamp'],
                    y=word_growth_df['word_change'],
                    name='Word Changes',
                    marker_color='rgb(140, 86, 75)'
                ),
                row=3, col=2
            )
            
            # Add more informative hover text for word changes
            fig.update_traces(
                hovertemplate="<br>".join([
                    "Date: %{x}",
                    "Change: %{y:+} words",
                    "<extra></extra>"
                ]),
                row=3, col=2
            )
        
        # Update layout
        fig.update_layout(
            height=1200 if word_growth_df is not None and not word_growth_df.empty else 800,
            title_text=title,
            title_x=0.5,
            showlegend=True,
            template='plotly_white',
            font=dict(size=12),
            legend=dict(
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=0.01
            )
        )
        
        # Update axes
        fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGrey')
        fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGrey')
        
        # Add helpful hover info for all traces
        fig.update_traces(
            hovertemplate="<br>".join([
                "Date: %{x}",
                "Count: %{y}",
                "<extra></extra>"
            ]),
            selector=dict(type='scatter')
        )
        
        try:
            fig.write_html('document_metrics.html')
            self.console.print("[green]Historical metrics plot saved as 'document_metrics.html'[/green]")
        except Exception as e:
            self.console.print(f"[red]Error saving plot: {str(e)}[/red]")

    def analyze_comments(self, comments: List[dict]) -> Dict[str, dict]:
        """Analyze comment activity."""
        comment_stats = defaultdict(lambda: {
            'comments_made': 0,
            'replies_made': 0,
            'resolved_comments': 0,
            'email': 'Email not available'
        })

        for comment in comments:
            author_info = comment.get('author', {})
            author = author_info.get('displayName', 'Unknown User')
            # Get email from the correct location in comment author structure
            email = author_info.get('me', False) and 'me' or author_info.get('email', 'Email not available')
            comment_stats[author]['comments_made'] += 1
            comment_stats[author]['email'] = email
            
            if comment.get('resolved', False):
                resolver_info = comment.get('resolvedBy', {})
                resolver = resolver_info.get('displayName', author)
                resolver_email = resolver_info.get('me', False) and 'me' or resolver_info.get('email', 'Email not available')
                comment_stats[resolver]['resolved_comments'] += 1
                comment_stats[resolver]['email'] = resolver_email
            
            for reply in comment.get('replies', []):
                reply_author_info = reply.get('author', {})
                reply_author = reply_author_info.get('displayName', 'Unknown User')
                reply_email = reply_author_info.get('me', False) and 'me' or reply_author_info.get('email', 'Email not available')
                comment_stats[reply_author]['replies_made'] += 1
                comment_stats[reply_author]['email'] = reply_email

        return dict(comment_stats)

    def count_words(self, text: str) -> int:
        """Count words in text, focusing on actual document content."""
        try:
            
            
            # Remove URLs and email addresses
            text = re.sub(r'http\S+|www\.\S+|\S+@\S+', '', text)
            
            # Remove special characters but keep sentence structure
            text = re.sub(r'[^\w\s.,!?"-]', ' ', text)
            
            # Remove comment sections and metadata
            patterns_to_remove = [
                r'Comments:.*?(?=\n\n|\Z)',  # Remove comment blocks
                r'Suggested edits:.*?(?=\n\n|\Z)',  # Remove suggestion blocks
                r'Last edited.*?(?=\n|\Z)',  # Remove edit information
                r'\[.*?\]',  # Remove bracketed content
                r'\{.*?\}'   # Remove curly brace content
            ]
            
            for pattern in patterns_to_remove:
                text = re.sub(pattern, '', text, flags=re.DOTALL | re.IGNORECASE)
            
            # Normalize whitespace
            text = ' '.join(text.split())
            
            # Count non-empty words
            words = [word for word in text.split() if word.strip() and not word.isdigit()]
            return len(words)
            
        except Exception as e:
            self.console.print(f"[red]Error counting words: {str(e)}[/red]")
            return 0
        
    
    def display_analytics(self, doc_id: str) -> None:
        """Display comprehensive analytics for the document."""
        try:
            # Header
            self.console.print("\n[bold blue]Document Analytics Report[/bold blue]")
            
            # 1. Document Metadata
            metadata = self._display_document_metadata(doc_id)
            
            # 2. Revision Analysis
            revisions = self.get_revision_history(doc_id)
            word_growth_df = None
            if revisions:
                word_growth_df = self._display_revision_analysis(doc_id, revisions)
            
            # 3. Activity Analysis
            activities = self._display_activity_analysis(doc_id)
            
            # 4. Comment Analysis
            comments = self._display_comment_analysis(doc_id)
            
            # 5. Historical Analysis and Summary
            if comments or activities:
                self._display_historical_analysis(
                    activities, 
                    comments, 
                    word_growth_df, 
                    metadata.get('title', 'Document')
                )
                
        except Exception as e:
            self.console.print(f"[red]Error generating analytics report: {str(e)}[/red]")

    def _display_document_metadata(self, doc_id: str) -> dict:
        """Display document metadata and return it for further use."""
        metadata = self.get_document_metadata(doc_id)
        if metadata:
            table = Table(title="\nDocument Information")
            table.add_column("Field")
            table.add_column("Value")
            
            for field, value in metadata.items():
                table.add_row(field.replace('_', ' ').title(), str(value))
            self.console.print(table)
        return metadata

    def analyze_word_growth(self, doc_id: str, revisions: List[dict]) -> pd.DataFrame:
        """Analyze the growth of total words over time and provide user statistics."""
        import time
        word_history = []
        prev_word_count = 0
        total_revisions = len(revisions)
        
        self.console.print(f"\n[yellow]Analyzing {total_revisions} revisions for word count history...[/yellow]")
        
        sorted_revisions = sorted(revisions, key=lambda x: x.get('modifiedTime', ''))
        
        for i, revision in enumerate(sorted_revisions, 1):
            revision_id = revision.get('id')
            mod_time = revision.get('modifiedTime')
            user_info = revision.get('lastModifyingUser', {})
            user = user_info.get('displayName', 'Unknown User')
            email = self._get_user_email(user_info)
            
            if revision_id and mod_time:
                self.console.print(f"Processing revision {i}/{total_revisions}")
                
                content = self.get_revision_content(doc_id, revision_id)
                word_count = self.count_words(content)
                
                word_diff = word_count - prev_word_count
                if abs(word_diff) >= 2 or i == 1 or i == total_revisions:
                    word_history.append({
                        'timestamp': datetime.fromisoformat(mod_time.replace('Z', '+00:00')),
                        'total_words': word_count,
                        'word_change': word_diff,
                        'user': user,
                        'email': email
                    })
                    prev_word_count = word_count
                
                if i < total_revisions:
                    time.sleep(3)

        self.console.print("[green]Word count analysis complete![/green]")
        
        df = pd.DataFrame(word_history)
        if not df.empty:
            df.sort_values('timestamp', inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            user_stats = df.groupby(['user', 'email']).agg({
                'word_change': [
                    ('total_words_added', lambda x: x[x > 0].sum()),
                    ('total_words_removed', lambda x: abs(x[x < 0].sum())),
                    ('net_words_added', 'sum'),
                    ('number_of_edits', 'count')
                ]
            })
            
            user_stats.columns = user_stats.columns.get_level_values(1)
            user_stats = user_stats.reset_index()
            
            user_stats['avg_words_per_edit'] = (user_stats['total_words_added'] + user_stats['total_words_removed']) / user_stats['number_of_edits']
            user_stats = user_stats.sort_values('net_words_added', ascending=False)
            
            stats_table = Table(title="\nUser Word Count Statistics")
            stats_table.add_column("User")
            stats_table.add_column("Email")
            stats_table.add_column("Words Added")
            stats_table.add_column("Words Removed")
            stats_table.add_column("Net Change")
            stats_table.add_column("Number of Edits")
            stats_table.add_column("Avg Words/Edit")
            
            for _, row in user_stats.iterrows():
                stats_table.add_row(
                    row['user'],
                    row['email'],
                    f"{int(row['total_words_added']):,}",
                    f"{int(row['total_words_removed']):,}",
                    f"{int(row['net_words_added']):,}",
                    str(row['number_of_edits']),
                    f"{row['avg_words_per_edit']:.1f}"
                )
            
            self.console.print(stats_table)
        
        return df

    def _display_revision_analysis(self, doc_id: str, revisions: List[dict]) -> Optional[pd.DataFrame]:
        try:
            word_growth_df = self.analyze_word_growth(doc_id, revisions)
            
            if not word_growth_df.empty:
                table = Table(title="\nWord Count History")
                table.add_column("Timestamp")
                table.add_column("Total Words")
                table.add_column("Word Change")
                table.add_column("User")
                table.add_column("Email")
                
                for _, row in word_growth_df.iterrows():
                    table.add_row(
                        row['timestamp'].strftime('%Y-%m-%d %H:%M:%S'),
                        str(row['total_words']),
                        f"{row['word_change']:+d}",
                        row['user'],
                        row['email']
                    )
                self.console.print(table)
            
            contributions = self.analyze_contributions(revisions)
            table = Table(title="\nUser Contributions")
            table.add_column("User")
            table.add_column("Email")
            table.add_column("Revisions")
            table.add_column("First Modified")
            table.add_column("Last Modified")

            for user, stats in sorted(contributions.items(), key=lambda x: x[1]['revision_count'], reverse=True):
                table.add_row(
                    user,
                    stats['email'],
                    str(stats['revision_count']),
                    str(stats['first_modified']),
                    str(stats['last_modified'])
                )
            self.console.print(table)
            
            return word_growth_df
        
        except Exception as e:
            self.console.print(f"[yellow]Error in revision analysis: {str(e)}[/yellow]")
            return None
    
    def _display_activity_analysis(self, doc_id: str) -> Optional[List[dict]]:
        """Display activity analysis and return activities for further use."""
        try:
            activities = self.get_activity_history(doc_id)
            # print(activities[0])
            if activities:
                word_contributions = self.calculate_word_contributions(activities)
                
                if word_contributions:
                    table = Table(title="\nEstimated Word Contributions")
                    table.add_column("User")
                    table.add_column("Estimated Words")
                    
                    for user, words in sorted(word_contributions.items(), key=lambda x: x[1], reverse=True):
                        table.add_row(user, str(words))
                    
                    self.console.print(table)
            return activities
            
        except Exception as e:
            self.console.print("[yellow]Word contribution analysis not available. Enable Drive Activity API for this feature.[/yellow]")
            return None

    def _display_comment_analysis(self, doc_id: str) -> Optional[List[dict]]:
        """Display comment analysis and return comments for further use."""
        comments = self.get_comments(doc_id)
        if comments:
            comment_stats = self.analyze_comments(comments)
            
            table = Table(title="\nComment Activity")
            table.add_column("User")
            table.add_column("Email")
            table.add_column("Comments Made")
            table.add_column("Replies Made")
            table.add_column("Comments Resolved")

            for user, stats in sorted(comment_stats.items(), key=lambda x: x[1]['comments_made'], reverse=True):
                table.add_row(
                    user,
                    stats['email'],
                    str(stats['comments_made']),
                    str(stats['replies_made']),
                    str(stats['resolved_comments'])
                )
            self.console.print(table)
        return comments
    
    def _display_historical_analysis(
        self, 
        activities: Optional[List[dict]], 
        comments: Optional[List[dict]], 
        word_growth_df: Optional[pd.DataFrame],
        doc_title: str
    ) -> None:
        """Display historical analysis and summary statistics."""
        try:
            df = self.create_historical_analysis(activities or [], comments or [])
            if not df.empty:
                self.plot_historical_metrics(df, f"Historical Metrics for {doc_title}", word_growth_df)
                
                # Summary statistics
                stats_table = Table(title="\nSummary Statistics")
                stats_table.add_column("Metric")
                stats_table.add_column("Value")
                
                # Activity metrics
                total_edits = df['edits'].sum()
                total_comments = df['comments'].sum()
                total_replies = df['replies'].sum()
                total_resolved = df['resolved'].sum()
                
                stats_table.add_row("Total Edits", str(total_edits))
                stats_table.add_row("Total Comments", str(total_comments))
                stats_table.add_row("Total Replies", str(total_replies))
                stats_table.add_row("Total Resolved Comments", str(total_resolved))
                
                if total_comments > 0:
                    resolution_rate = (total_resolved / total_comments) * 100
                    stats_table.add_row("Comment Resolution Rate", f"{resolution_rate:.1f}%")
                
                # Word count statistics
                if word_growth_df is not None and not word_growth_df.empty:
                    latest_word_count = word_growth_df['total_words'].iloc[-1]
                    total_word_changes = word_growth_df['word_change'].abs().sum()
                    avg_words_per_edit = total_word_changes / len(word_growth_df) if len(word_growth_df) > 0 else 0
                    
                    stats_table.add_row("Current Word Count", str(latest_word_count))
                    stats_table.add_row("Total Word Changes", str(total_word_changes))
                    stats_table.add_row("Average Words per Edit", f"{avg_words_per_edit:.1f}")
                
                self.console.print(stats_table)
                
        except Exception as e:
            self.console.print(f"[yellow]Error in historical analysis: {str(e)}[/yellow]")    
        
def main():
    parser = argparse.ArgumentParser(description='Google Docs Analytics Tool')
    parser.add_argument('doc_id', help='The ID of the Google Doc to analyze')
    args = parser.parse_args()

    analyzer = GoogleDocsAnalyzer()
    
    try:
        analyzer.authenticate()
        analyzer.display_analytics(args.doc_id)
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()