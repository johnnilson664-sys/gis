import json, os, subprocess, sys, tarfile, termios
from pathlib import Path

root = Path('/workspace/work/ti-downloads')
archive = Path('/workspace/work/ti-downloads-deploy.tar.gz')
settings = termios.tcgetattr(sys.stdin.fileno())
hidden = settings[:]
hidden[3] &= ~termios.ECHO
termios.tcsetattr(sys.stdin.fileno(), termios.TCSANOW, hidden)
print('Ready for source workflow JSON on stdin (input is hidden).', flush=True)
try:
    payload = json.loads(sys.stdin.readline())
finally:
    termios.tcsetattr(sys.stdin.fileno(), termios.TCSANOW, settings)
credential = payload['credential']
assert credential['auth_mode'] == 'http_extra_header'
manifest = json.loads((root/'.openai/hosting.json').read_text())
assert manifest['project_id'] == payload['project_id']
env = os.environ.copy()
env['GIT_TERMINAL_PROMPT'] = '0'
count = int(env.get('GIT_CONFIG_COUNT', '0'))
env['GIT_CONFIG_COUNT'] = str(count + 1)
env[f'GIT_CONFIG_KEY_{count}'] = 'http.extraHeader'
env[f'GIT_CONFIG_VALUE_{count}'] = 'Authorization: Bearer ' + credential['token']
def git(args):
    result = subprocess.run(['git'] + args, cwd=root, env=env, text=True, capture_output=True)
    if result.returncode:
        message = result.stderr.replace(credential['token'], '[redacted]')
        raise RuntimeError(message)
    return result.stdout.strip()

branch = credential['branch']
git(['init','-b',branch])
git(['config','user.name','Codex'])
git(['config','user.email','codex@localhost'])
remotes = git(['remote']).splitlines()
if 'origin' not in remotes:
    git(['remote','add','origin',credential['remote_url']])
else:
    assert git(['remote','get-url','origin']) == credential['remote_url']
heads = git(['ls-remote','--heads','origin',f'refs/heads/{branch}'])
if heads and not (root/'.git/refs/heads'/branch).exists():
    git(['fetch','origin',branch])
    git(['reset','--mixed','FETCH_HEAD'])
git(['add','dist','.openai/hosting.json'])
if git(['status','--porcelain']):
    git(['commit','-m','Add standalone 3CX talk and downloadable materials'])
sha = git(['rev-parse','HEAD'])
git(['push','origin',f'HEAD:refs/heads/{branch}'])
remote_sha = git(['ls-remote','--heads','origin',f'refs/heads/{branch}']).split()[0]
assert remote_sha == sha
with tarfile.open(archive,'w:gz') as output:
    output.add(root/'.openai/hosting.json',arcname='.openai/hosting.json')
    output.add(root/'dist',arcname='dist')
result = {'project_id':manifest['project_id'],'commit_sha':sha,'archive':str(archive)}
(Path('/workspace/work')/'ti-downloads-source-result.json').write_text(json.dumps(result,indent=2))
print(json.dumps(result),flush=True)
