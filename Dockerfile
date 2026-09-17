FROM node:24-alpine

WORKDIR /app

COPY package*.json ./
RUN npm install 

COPY . .

EXPOSE 5173

CMD ["npm", "run", "playground", "--", "--host", "0.0.0.0"]