# Use the latest Node.js LTS version
FROM node:20-alpine

# Create app directory
WORKDIR /usr/src/app

# Install app dependencies
COPY package*.json ./
RUN npm install

# Copy app source code
COPY . .

# Expose the port your app uses (assumed 3000; adjust if different)
EXPOSE 3000

# Start the app in development mode with nodemon
CMD ["npm", "run", "devStart"]
# Use nodemon for hot reloading during development