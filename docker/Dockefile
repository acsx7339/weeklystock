FROM python:3.9.20

# Install system dependencies and TA-Lib
RUN apt-get update && \
    apt-get install -y build-essential wget && \
    wget http://prdownloads.sourceforge.net/ta-lib/ta-lib-0.4.0-src.tar.gz && \
    tar -xzf ta-lib-0.4.0-src.tar.gz && \
    cd ta-lib-0.4.0 && \
    ./configure --prefix=/usr && \
    make && \
    make install && \
    cd .. && \
    rm -rf ta-lib-0.4.0 ta-lib-0.4.0-src.tar.gz

# Install Python dependencies
RUN pip install --upgrade pip && \
    pip install twstock FinMind ta-lib yfinance pandas numpy scipy tqdm openpyxl

# Create a user and switch to non-root user
RUN adduser --disabled-password --gecos "" simon && \
    echo "simon:simon" | chpasswd && \
    adduser simon sudo

USER simon

# Set the working directory (optional, for your application)
WORKDIR /home/simon

