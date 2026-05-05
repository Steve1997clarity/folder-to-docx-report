# Use an official Python slim image.
FROM python:3.12-slim

# Set environment variables.
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set working directory.
WORKDIR /app

# Copy and install dependencies.
COPY requirements.txt /app/
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copy the project files into the container.
COPY . /app/

# Expose the port.
EXPOSE 11312

# Run the Flask app.
CMD ["python", "main.py"]
