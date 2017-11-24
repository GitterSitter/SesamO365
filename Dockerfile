FROM node:8
MAINTAINER Trond "trond.tufte@bouvet.no"
ADD . /node-office 
WORKDIR /node-office
COPY package.json /src
RUN npm install
EXPOSE 8000
CMD [ "npm", "start" ]
