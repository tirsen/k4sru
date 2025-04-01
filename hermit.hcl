version = "1"

channel "python" {
  url = "https://github.com/cashapp/hermit-packages/releases/download/python-3.11.7/python-3.11.7.tar.gz"
  sha256 = "sha256-0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0"
}

packages {
  python {
    version = "3.11.7"
    channel = "python"
  }
}

environments {
  default {
    packages = ["python"]
    activation = "source .hermit/python/bin/activate"
  }
}

hooks {
  install = [
    "pip install -r requirements.txt"
  ]
} 