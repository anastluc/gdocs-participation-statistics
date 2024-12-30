import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from collections import defaultdict
from datetime import datetime, timedelta
import pickle
import argparse
from typing import Dict, List, Tuple
from rich.console import Console
from rich.table import Table
import pandas as pd
from plotly import graph_objects as go
from plotly.subplots import make_subplots

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
        self.creds = None
        self.service = None
        self.drive_service = None
        self.activity_service = None
        self.console = Console()

    def authenticate(self):
        """Handle Google OAuth authentication."""
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)
        
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
            
            with open('token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

        self.service = build('docs', 'v1', credentials=self.creds)
        self.drive_service = build('drive', 'v3', credentials=self.creds)
        self.activity_service = build('driveactivity', 'v2', credentials=self.creds)
        
        # Verify authentication
        try:
            about = self.drive_service.about().get(fields="user").execute()
            self.console.print(f"[green]Successfully authenticated as: {about['user']['emailAddress']}[/green]")
        except Exception as e:
            self.console.print(f"[red]Authentication verification failed: {str(e)}[/red]")

    def get_document_metadata(self, doc_id: str) -> dict:
        """Retrieve basic document metadata."""
        try:
            document = self.service.documents().get(documentId=doc_id).execute()
            file_metadata = self.drive_service.files().get(
                fileId=doc_id, 
                fields="createdTime,modifiedTime,owners,lastModifyingUser"
            ).execute()
            
            return {
                'title': document.get('title', 'Untitled'),
                'created_time': file_metadata.get('createdTime', 'Unknown'),
                'modified_time': file_metadata.get('modifiedTime', 'Unknown'),
                'owner': file_metadata.get('owners', [{}])[0].get('displayName', 'Unknown'),
                'last_modifier': file_metadata.get('lastModifyingUser', {}).get('displayName', 'Unknown')
            }
        except Exception as e:
            self.console.print(f"[red]Error retrieving document metadata: {str(e)}[/red]")
            return {}

    def get_revision_history(self, doc_id: str) -> List[dict]:
        """Get the complete revision history of the document."""
        try:
            self.console.print("[yellow]Checking revision access...[/yellow]")
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

    def get_activity_history(self, doc_id: str) -> List[dict]:
        """Get detailed activity history including content changes."""
        try:
            activities = []
            page_token = None
            
            while True:
                results = self.activity_service.activity().query(
                    body={
                        'itemName': f'items/{doc_id}',
                        'pageSize': 100,
                        'pageToken': page_token
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
            self.console.print("[yellow]Checking comment access...[/yellow]")
            comments = []
            page_token = None
            
            while True:
                response = self.drive_service.comments().list(
                    fileId=doc_id,
                    fields="nextPageToken,comments(id,content,author,createdTime,resolved,replies)",
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

    def analyze_contributions(self, revisions: List[dict]) -> Dict[str, dict]:
        """Analyze user contributions from revision history."""
        contributions = defaultdict(lambda: {
            'revision_count': 0,
            'last_modified': None,
            'first_modified': None
        })

        for revision in revisions:
            user = revision.get('lastModifyingUser', {}).get('displayName', 'Unknown User')
            mod_time = revision.get('modifiedTime')
            
            contributions[user]['revision_count'] += 1
            
            if not contributions[user]['last_modified'] or mod_time > contributions[user]['last_modified']:
                contributions[user]['last_modified'] = mod_time
            
            if not contributions[user]['first_modified'] or mod_time < contributions[user]['first_modified']:
                contributions[user]['first_modified'] = mod_time

        return dict(contributions)

    def calculate_word_contributions(self, doc_id: str, activities: List[dict]) -> Dict[str, int]:
        """Calculate approximate word contributions per user based on activity history."""
        word_counts = defaultdict(int)
        
        try:
            for activity in activities:
                if 'target' in activity and 'driveItem' in activity['target']:
                    actions = activity.get('primaryActionDetail', {})
                    if 'edit' in actions:
                        actor = self.get_actor_name(activity)
                        word_counts[actor] += 1
            
            # Normalize the counts (multiply by average words per edit)
            avg_words_per_edit = 5  # This is an estimate
            for user in word_counts:
                word_counts[user] *= avg_words_per_edit
                
        except Exception as e:
            self.console.print(f"[red]Error calculating word contributions: {str(e)}[/red]")
        
        return dict(word_counts)

    def analyze_comments(self, comments: List[dict]) -> Dict[str, dict]:
        """Analyze comment activity."""
        comment_stats = defaultdict(lambda: {
            'comments_made': 0,
            'replies_made': 0,
            'resolved_comments': 0
        })

        for comment in comments:
            author = comment.get('author', {}).get('displayName', 'Unknown User')
            comment_stats[author]['comments_made'] += 1
            
            if comment.get('resolved', False):
                resolver = comment.get('resolvedBy', {}).get('displayName', author)
                comment_stats[resolver]['resolved_comments'] += 1
            
            for reply in comment.get('replies', []):
                reply_author = reply.get('author', {}).get('displayName', 'Unknown User')
                comment_stats[reply_author]['replies_made'] += 1

        return dict(comment_stats)

    def get_actor_name(self, activity: dict) -> str:
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
        
        # Process activities
        for activity in activities:
            timestamp = activity.get('timestamp', '')
            if timestamp:
                date = datetime.fromisoformat(timestamp.replace('Z', '+00:00')).date()
                metrics[date]['edits'] += 1
                
        # Process comments
        for comment in comments:
            created_time = comment.get('createdTime', '')
            if created_time:
                date = datetime.fromisoformat(created_time.replace('Z', '+00:00')).date()
                metrics[date]['comments'] += 1
                
                # Count replies
                for reply in comment.get('replies', []):
                    reply_time = reply.get('createdTime', '')
                    if reply_time:
                        reply_date = datetime.fromisoformat(reply_time.replace('Z', '+00:00')).date()
                        metrics[reply_date]['replies'] += 1
                
                # Count resolutions
                if comment.get('resolved', False):
                    resolved_time = comment.get('resolvedTime', created_time)
                    resolved_date = datetime.fromisoformat(resolved_time.replace('Z', '+00:00')).date()
                    metrics[resolved_date]['resolved'] += 1
        
        # Convert to DataFrame
        dates = sorted(metrics.keys())
        data = {
            'date': dates,
            'edits': [metrics[date]['edits'] for date in dates],
            'comments': [metrics[date]['comments'] for date in dates],
            'replies': [metrics[date]['replies'] for date in dates],
            'resolved': [metrics[date]['resolved'] for date in dates]
        }
        
        return pd.DataFrame(data)

    def plot_historical_metrics(self, df: pd.DataFrame, title: str):
        """Create and save an interactive plot of historical metrics."""
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=('Document Edits', 'Comments Made', 'Replies Made', 'Comments Resolved')
        )
        
        fig.add_trace(
            go.Scatter(x=df['date'], y=df['edits'], mode='lines+markers', name='Edits'),
            row=1, col=1
        )
        fig.add_trace(
            go.Scatter(x=df['date'], y=df['comments'], mode='lines+markers', name='Comments'),
            row=1, col=2
        )
        fig.add_trace(
            go.Scatter(x=df['date'], y=df['replies'], mode='lines+markers', name='Replies'),
            row=2, col=1
        )
        fig.add_trace(
            go.Scatter(x=df['date'], y=df['resolved'], mode='lines+markers', name='Resolved'),
            row=2, col=2
        )
        
        fig.update_layout(
            height=800,
            title_text=title,
            showlegend=True
        )
        
        fig.write_html('document_metrics.html')
        self.console.print("[green]Historical metrics plot saved as 'document_metrics.html'[/green]")

    def get_document_content(self, doc_id: str) -> dict:
        """Get the full document content."""
        try:
            return self.service.documents().get(documentId=doc_id).execute()
        except Exception as e:
            self.console.print(f"[red]Error retrieving document content: {str(e)}[/red]")
            return {}

    def display_analytics(self, doc_id: str):
        """Display comprehensive analytics for the document."""
        self.console.print("\n[bold blue]Document Analytics Report[/bold blue]")
        
        # Get document metadata
        metadata = self.get_document_metadata(doc_id)
        if metadata:
            self.console.print("\n[bold]Document Information:[/bold]")
            self.console.print(f"Title: {metadata['title']}")
            self.console.print(f"Created: {metadata['created_time']}")
            self.console.print(f"Last Modified: {metadata['modified_time']}")
            self.console.print(f"Owner: {metadata['owner']}")
            self.console.print(f"Last Modified By: {metadata['last_modifier']}")
        
        # Get revision history and analyze user contributions
        revisions = self.get_revision_history(doc_id)
        if revisions:
            contributions = self.analyze_contributions(revisions)
            
            contrib_table = Table(title="\nUser Contributions")
            contrib_table.add_column("User")
            contrib_table.add_column("Revisions")
            contrib_table.add_column("First Modified")
            contrib_table.add_column("Last Modified")

            for user, stats in contributions.items():
                contrib_table.add_row(
                    user,
                    str(stats['revision_count']),
                    stats['first_modified'],
                    stats['last_modified']
                )
            self.console.print(contrib_table)
        
        # Try to get activities and word contributions
        try:
            activities = self.get_activity_history(doc_id)
            if activities:
                word_contributions = self.calculate_word_contributions(doc_id, activities)
                
                word_table = Table(title="\nEstimated Word Contributions")
                word_table.add_column("User")
                word_table.add_column("Estimated Words")
                
                for user, words in word_contributions.items():
                    word_table.add_row(user, str(words))
                
                self.console.print(word_table)
        except Exception as e:
            self.console.print("[yellow]Word contribution analysis not available. Enable Drive Activity API for this feature.[/yellow]")
            activities = []

        # Get comments and analyze
        comments = self.get_comments(doc_id)
        if comments:
            comment_stats = self.analyze_comments(comments)
            
            comment_table = Table(title="\nComment Activity")
            comment_table.add_column("User")
            comment_table.add_column("Comments Made")
            comment_table.add_column("Replies Made")
            comment_table.add_column("Comments Resolved")

            for user, stats in comment_stats.items():
                comment_table.add_row(
                    user,
                    str(stats['comments_made']),
                    str(stats['replies_made']),
                    str(stats['resolved_comments'])
                )
            self.console.print(comment_table)

        
        self.console.print("\n[bold blue]Document Analytics Report[/bold blue]")
        
        # Get document metadata
        document = self.get_document_content(doc_id)
        
        # Try to get activities and word contributions if Drive Activity API is available
        try:
            activities = self.get_activity_history(doc_id)
            if activities:
                word_contributions = self.calculate_word_contributions(doc_id, activities)
                
                # Create word contributions table
                word_table = Table(title="\nEstimated Word Contributions")
                word_table.add_column("User")
                word_table.add_column("Estimated Words")
                
                for user, words in word_contributions.items():
                    word_table.add_row(user, str(words))
                
                self.console.print(word_table)
        except Exception as e:
            self.console.print("[yellow]Word contribution analysis not available. Enable Drive Activity API for this feature.[/yellow]")
            activities = []
        
        # Get comments for historical analysis
        comments = self.get_comments(doc_id)
        
        # Create and display historical analysis
        try:
            df = self.create_historical_analysis(activities, comments)
            self.plot_historical_metrics(df, f"Historical Metrics for {document.get('title', 'Document')}")
            
            # Display summary statistics
            total_edits = df['edits'].sum()
            total_comments = df['comments'].sum()
            total_replies = df['replies'].sum()
            total_resolved = df['resolved'].sum()
            
            stats_table = Table(title="\nSummary Statistics")
            stats_table.add_column("Metric")
            stats_table.add_column("Value")
            
            stats_table.add_row("Total Edits", str(total_edits))
            stats_table.add_row("Total Comments", str(total_comments))
            stats_table.add_row("Total Replies", str(total_replies))
            stats_table.add_row("Total Resolved Comments", str(total_resolved))
            
            self.console.print(stats_table)
        except Exception as e:
            self.console.print("[yellow]Historical analysis not available. Check if Drive Activity API is enabled.[/yellow]")
            
            # Still show comment statistics if available
            if comments:
                comment_stats = self.analyze_comments(comments)
                stats_table = Table(title="\nComment Statistics")
                stats_table.add_column("User")
                stats_table.add_column("Comments Made")
                stats_table.add_column("Replies Made")
                stats_table.add_column("Comments Resolved")
                
                for user, stats in comment_stats.items():
                    stats_table.add_row(
                        user,
                        str(stats['comments_made']),
                        str(stats['replies_made']),
                        str(stats['resolved_comments'])
                    )
                self.console.print(stats_table)
    
    
    

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