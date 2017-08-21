FROM node:6
MAINTAINER Trond "trond.tufte@bouvet.no"
ADD . /node-office 
WORKDIR /node-office
RUN npm install
EXPOSE 8000
CMD [ "npm", "start" ]
