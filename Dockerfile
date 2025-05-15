# Use an official Python runtime as a parent image
FROM python:3.12

# Set the working directory in the container
WORKDIR /Demo-Deployment

# Copy the current directory contents into the container at /app
COPY . /Demo-Deployment

# Install required Python packages
RUN pip install --no-cache-dir -r requirements.txt

# Expose the port Streamlit runs on
EXPOSE 8501

# Command to run the Streamlit app
CMD ["streamlit", "run", "frontend.py", "--server.port=8501", "--server.address=0.0.0.0"]

