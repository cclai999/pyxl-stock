import requests


def get_html_to_file(url:str, fname: str):
    resp = requests.get(url)
    resp.raise_for_status()
    f = open(fname, "w")
    f.write(resp.text)
    f.close()