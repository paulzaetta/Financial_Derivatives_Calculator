# Structure du dossier

## Évaluation_Produits_Dérivés.pdf

Ce fichier (en format PDF) contient une explication détaillée du code VBA "vba_project.xlsm".

## vba_project.xlsm

Ce projet porte sur la création d'un outil permettant d'évaluer certains produits dérivés tels que des options, des obligations et des swpas. L'outil développé se décompose en quatre applications.

Une première application a été développé permettant à l'utilisateur de pouvoir valoriser de nombreuses options, telles que les options dites "Vanilles" à l'aide du modèle de Black et Scholes ou du modèle Binomial (options européennes ou américaines). Il est également possible d'évaluer des options exotiques telles que les options à barrières, les options asiatiques et les options binaires. La volatilité implicite peut également être calculée mais uniquement sous certaines conditions. L'évaluation du prix de l'option est accompagnée par l'évaluation de ses "grecques" (delta, gamma, theta, vega, rho). 

Une seconde application permet de réaliser des simulations de Monte Carlo. Ces simulations permettent de valoriser des options européennes selon le modèle log-normal et le modèle de diffusion à sauts de Merton. 

Une troisième application évalue le prix des obligations selon la périodicité des versements des intérêts et des autres propriétés fixées par l'utilisateur. Cette application calcule également d'autres caractéristiques telles que le taux de rendement actuariel, la duration, la sensibilité ainsi que la convexité. 

Enfin, une dernière application a été développé permettant de calculer la valeur d'un swap.
