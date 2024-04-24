import os

from flask import Flask, render_template, request, redirect, url_for, make_response
from github import Github
import xlsxwriter
import io

token = os.getenv('token')
g = Github(base_url="https://api.github.com", login_or_token=token)
repo = g.get_repo("pylearn-devops/ab-github-actions")

app = Flask(__name__, template_folder="templates")


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/ready-for-review', methods=['GET', 'POST'])
def ready_for_review():
    if request.method == 'POST' and request.form.get('action') == 'refresh':
        pull_requests = fetch_ready_for_review_pull_requests()
        return render_template('ready_for_review.html', ready_for_review_prs=pull_requests)
    elif request.method == 'GET' and request.args.get('download') == 'excel':
        pull_requests = fetch_ready_for_review_pull_requests()
        excel_data = generate_excel_data(pull_requests)
        return send_excel_file(excel_data, 'ready_for_review_pull_requests.xlsx')
    else:
        pull_requests = fetch_ready_for_review_pull_requests()
        return render_template('ready_for_review.html', ready_for_review_prs=pull_requests)


@app.route('/ready-for-release', methods=['GET', 'POST'])
def ready_for_release():
    if request.method == 'POST' and request.form.get('action') == 'refresh':
        pull_requests = fetch_ready_for_release_pull_requests()
        return render_template('ready_for_release.html', ready_for_release_prs=pull_requests)
    else:
        pull_requests = fetch_ready_for_release_pull_requests()
        return render_template('ready_for_release.html', ready_for_release_prs=pull_requests)


def fetch_ready_for_review_pull_requests():
    # Get the repository

    # Fetch all open pull requests
    pull_requests = repo.get_pulls(state="open")

    # Filter pull requests with the "Ready for Review" label
    ready_for_review_prs = [pr for pr in pull_requests if
                            "ready for review" in [label.name for label in pr.get_labels()]]
    return ready_for_review_prs


def fetch_ready_for_release_pull_requests():
    # Get the repository

    # Fetch all open pull requests
    pull_requests = repo.get_pulls(state="open")

    # Filter pull requests with the "Ready for Release" label
    ready_for_release_prs = [pr for pr in pull_requests if
                             "ready for release" in [label.name for label in pr.get_labels()]]
    return ready_for_release_prs


def generate_excel_data(pull_requests):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    # Write headers
    headers = ['Number', 'Title', 'Pull Request Link', 'Author']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # Write pull request data
    for row, pr in enumerate(pull_requests, start=1):
        worksheet.write(row, 0, pr.number)
        worksheet.write(row, 1, pr.title)
        worksheet.write(row, 2, pr.html_url)
        worksheet.write(row, 3, pr.user.login)

    workbook.close()
    output.seek(0)
    return output.getvalue()


def send_excel_file(data, filename):
    response = make_response(data)
    response.headers['Content-Disposition'] = f'attachment; filename={filename}'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response


@app.route('/releases')
def releases():
    # Fetch all tags
    tags = repo.get_tags()

    # Initialize a list to store tag data
    tag_data = []

    # Iterate over tags and fetch associated commits
    for tag in tags:
        # Get the commit associated with the tag
        commit = tag.commit

        # Get the commit message
        commit_message = commit.commit.message

        # Get the commit author
        commit_author = commit.author.login if commit.author else "Unknown"

        # Append tag data to the list
        tag_data.append({
            "tag_name": tag.name,
            "commit_message": commit_message,
            "commit_author": commit_author
        })

    return render_template('releases.html', tag_data=tag_data)


@app.route('/ready-for-reviews')
def ready_for_reviews():
    # Get the organization
    org = g.get_organization('pylearn-devops')

    # Fetch all repositories in the organization
    repos = org.get_repos()

    # Initialize a list to store repository names
    repo_names = []

    # Iterate over repositories
    for r in repos:
        # Fetch pull requests with "ready for review" label
        pull_requests = r.get_pulls(state='open', sort='created', base='master')

        # Filter pull requests with "ready for review" label
        ready_for_review_prs = [pr for pr in pull_requests if 'ready for review' in [label.name for label in pr.labels]]

        # If there are pull requests ready for review, add the repository name to the list
        if ready_for_review_prs:
            repo_names.append(r.name)

    return render_template('ready_for_reviews.html', repo_names=repo_names)


@app.route('/ready-for-review/<repo_name>')
def repo_pull_requests(repo_name):
    # Get the organization
    org = g.get_organization('pylearn-devops')

    # Get the repository by name
    r = org.get_repo(repo_name)

    # Fetch pull requests with "ready for review" label
    pull_requests = r.get_pulls(state='open', sort='created', base='master')

    # Filter pull requests with "ready for review" label
    ready_for_review_prs = [pr for pr in pull_requests if 'ready for review' in [label.name for label in pr.labels]]

    # Render pull requests template with pull requests data
    return render_template('repo_pull_requests.html', repo_name=repo_name, pull_requests=ready_for_review_prs)


if __name__ == '__main__':
    app.run(debug=True)
