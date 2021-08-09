SELECT *
FROM PortfolioProject.dbo.CovidDeaths$
WHERE continent is Not NULL
ORDER BY 3, 4

--SELECT *
--FROM CovidVaccinations
--ORDER BY 3, 4

--SELECT  Location, date, total_cases, new_cases, total_deaths, population
--FROM PortfolioProject.dbo.CovidDeaths$
--ORDER BY 1, 2

-- Looking at Total Cases vs Total Deaths
-- Shows likelihood of dying if you contract covid
SELECT  Location, date, total_cases, total_deaths, (total_deaths/total_cases)*100 AS 'DeathPercentage'
FROM PortfolioProject.dbo.CovidDeaths$
WHERE continent is Not NULL
ORDER BY 1, 2


-- Looking at total cases vs population
-- Shows what percentage of population got covid

SELECT  Location, date, population, total_cases, (total_cases/population)*100 AS 'InfectionPercentage'
FROM PortfolioProject.dbo.CovidDeaths$
--WHERE location like '%states%'
WHERE continent is Not NULL
ORDER BY 1, 2


-- Looking at Countries with Highest Infection Rate compared to Population

SELECT Location, population, MAX(total_cases) AS HighestInfectionCount, MAX((total_cases/population))*100 AS 'PercentPopulationInfected'
FROM PortfolioProject.dbo.CovidDeaths$
--WHERE location like '%states%'
WHERE continent is Not NULL
GROUP BY Location, population
ORDER BY 4 DESC

--Showing Countries with Highest Death Count per Population

SELECT Location, MAX(cast(Total_deaths as int)) AS TotalDeathCount
FROM PortfolioProject.dbo.CovidDeaths$
WHERE continent is Not NULL
GROUP BY Location
ORDER BY 2 DESC


-- LET'S BREAK THINGS DOWN BY CONTINENT
--Showing continents with the highest death count per population
SELECT location, MAX(cast(Total_deaths as int)) AS TotalDeathCount
FROM PortfolioProject.dbo.CovidDeaths$
WHERE continent is NULL
GROUP BY location
ORDER BY 2 DESC



-- Global Numbers

SELECT date, SUM(new_cases) AS total_cases, SUM(cast(new_deaths as int)) as total_deaths, SUM(cast(new_deaths as int))/SUM(new_cases)*100 AS DeathPercentage
FROM PortfolioProject.dbo.CovidDeaths$
WHERE continent is Not NULL
GROUP BY date
ORDER BY 1, 2

-- Total Global
SELECT SUM(new_cases) AS total_cases, SUM(cast(new_deaths as int)) as total_deaths, SUM(cast(new_deaths as int))/SUM(new_cases)*100 AS DeathPercentage
FROM PortfolioProject.dbo.CovidDeaths$
WHERE continent is Not NULL
ORDER BY 1, 2


-- Looking at Total Population vs Vaccinations

--SELECT dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations
--, SUM(CONVERT(int, vac.new_vaccinations)) OVER (PARTITION BY dea.location ORDER BY dea.location, dea.date)
-- AS RollingPeopleVaccinated
--, (
--FROM PortfolioProject..CovidDeaths$ dea
--JOIN PortfolioProject..CovidVaccinations$ vac
--	ON dea.location = vac.location
--	AND dea.date = vac.date
--WHERE dea.continent is Not NULL
--ORDER BY 2, 3

-- USE CTE

WITH PopvsVac (Continent, location, date, population, new_vaccinations, RollingPeopleVaccinated)
AS 
(
SELECT dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations
, SUM(CONVERT(int, vac.new_vaccinations)) OVER (PARTITION BY dea.location ORDER BY dea.location, dea.date)
 AS RollingPeopleVaccinated
FROM PortfolioProject..CovidDeaths$ dea
JOIN PortfolioProject..CovidVaccinations$ vac
	ON dea.location = vac.location
	AND dea.date = vac.date
WHERE dea.continent is Not NULL
)
SELECT *, (RollingPeopleVaccinated/Population)*100
FROM PopvsVac


-- Temp Table

DROP TABLE IF EXISTS #PercentPopulationVaccinated
CREATE TABLE #PercentPopulationVaccinated
(
Continent nvarchar(255),
Location nvarchar(255),
Date datetime,
Population numeric,
New_vaccinations numeric,
RollingPeopleVaccinated numeric
)

Insert into #PercentPopulationVaccinated
SELECT dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations
, SUM(CONVERT(int, vac.new_vaccinations)) OVER (PARTITION BY dea.location ORDER BY dea.location, dea.date)
 AS RollingPeopleVaccinated
FROM PortfolioProject..CovidDeaths$ dea
JOIN PortfolioProject..CovidVaccinations$ vac
	ON dea.location = vac.location
	AND dea.date = vac.date
WHERE dea.continent is Not NULL

SELECT *, (RollingPeopleVaccinated/Population)*100
FROM #PercentPopulationVaccinated



CREATE VIEW PercentPopulationVaccinated AS
SELECT dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations
, SUM(CONVERT(int, vac.new_vaccinations)) OVER (PARTITION BY dea.location ORDER BY dea.location, dea.date)
 AS RollingPeopleVaccinated
FROM PortfolioProject..CovidDeaths$ dea
JOIN PortfolioProject..CovidVaccinations$ vac
	ON dea.location = vac.location
	AND dea.date = vac.date
WHERE dea.continent is Not NULL



SELECT *
FROM PercentPopulationVaccinated