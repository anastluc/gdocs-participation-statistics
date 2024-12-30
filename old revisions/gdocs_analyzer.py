import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from collections import defaultdict
from datetime import datetime
import pickle
import argparse
from typing import Dict, List
from rich.console import Console
from rich.table import Table

# Full set of required scopes
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
        self.console = Console()

    def authenticate(self):
        """Handle Google OAuth authentication."""
        # Delete token if it exists to force re-authentication
        if os.path.exists('token.pickle'):
            os.remove('token.pickle')
            
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
        
        # Verify authentication and permissions
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
            # First verify if we have permission to access revisions
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
            self.console.print("[yellow]Note: Make sure you have edit access to the document[/yellow]")
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
            self.console.print("[yellow]Note: Make sure you have comment access to the document[/yellow]")
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

    def display_analytics(self, doc_id: str):
        """Display comprehensive analytics for the document."""
        self.console.print("\n[bold blue]Document Analytics Report[/bold blue]")
        
        # Get and display document metadata
        metadata = self.get_document_metadata(doc_id)
        if metadata:
            self.console.print("\n[bold]Document Information:[/bold]")
            self.console.print(f"Title: {metadata['title']}")
            self.console.print(f"Created: {metadata['created_time']}")
            self.console.print(f"Last Modified: {metadata['modified_time']}")
            self.console.print(f"Owner: {metadata['owner']}")
            self.console.print(f"Last Modified By: {metadata['last_modifier']}")
        
        # Get revision history and analyze
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