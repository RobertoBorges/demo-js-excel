/* eslint-disable @typescript-eslint/no-unused-vars */
/* global console setInterval, clearInterval */
import fetch from "node-fetch";

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

/**
 * Multiply two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The product of the two numbers.
 */
function multiply(first, second) {
  return first * second;
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Fetches data from a GitHub repository and returns it as a string array.
 * @customfunction
 * @param {string} url The URL of the GitHub repository API.
 * @returns {string[][]} The repository data.
 */
async function fetchGitHubData() {
  const url = "https://api.github.com/repos/robertoborges/demo-js-excel";
  try {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error("Network response was not ok " + response.statusText);
    }
    const data = await response.json();
    return [
      ["ID", "Name", "Full Name", "Description", "Stars", "Forks", "Open Issues"],
      [
        data.id,
        data.name,
        data.full_name,
        data.description,
        data.stargazers_count,
        data.forks_count,
        data.open_issues_count,
      ],
    ];
  } catch (error) {
    console.error("Fetch error: ", error);
    return [["Error fetching data"]];
  }
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param {string} message String to write.
 * @returns String to write.
 */
function logMessage(message) {
  console.log(message);

  return message;
}
