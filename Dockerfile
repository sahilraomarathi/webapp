FROM node:18
RUN apt-get update && apt-get install -y libreoffice python3 python3-pip
RUN pip3 install python-docx --break-system-packages
WORKDIR /app
COPY . .
RUN npm install
CMD ["node", "index.js"]
