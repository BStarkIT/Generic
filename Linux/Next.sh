Next
    10.0.0.200
    Bonded RR
    Kosh
    Local 240 SSD
        \local\cache
        \docker
        \local\local
        \local\next0    320
        \local\next1    1000
        \local\next2    1000
    docker  master
        portainer 
        tdarr
        handbrake

Cube
    10.0.0.201
    Bonded RR
    Kosh
    Local 500
        \docker
        \local\local
        \local\cube0    1000
        \local\cube1    2000
        \local\cube2    3000
    docker
        portainer
        tdarr client
        Plex
        omni
        nginx
        sonarr
        radarr
        lidarr


sudo apt-get install samba -y
        sudo apt-get install cifs-utils -y

 mkdir -p ~/.docker/cli-plugins/
 curl -SL https://github.com/docker/compose/releases/download/v2.24.5/docker-compose-linux-x86_64 -o ~/.docker/cli-plugins/docker-compose
 chmod +x ~/.docker/cli-plugins/docker-compose
 docker compose version
 
Docker token dckr_pat_5DJ9lAr5RAGbj15GbPuQHkLMez4

Plex-Token=Jok_i15ixsrarw7Tz_KM

docker run -d -p 8000:8000 -p 9443:9443 --name=portainer --restart=always -v /var/run/docker.sock:/var/run/docker.sock -v portainer_data:/data portainer/portainer-ee:latest


Rack
    Local
    0       3TB     Music, Porn, ebooks     Done
    1       4TB     Syfy                    Done
    2       4TB     Films                   Done
    3       4TB     Films 2                 Done
    4       5TB     Active                  Done

Cube
    Local
    0       1TB     Documents           Done
    1       2TB     Classic             Done
    2       3TB     
    3       2TB     Anime               Done
    4       1TB     
    5       2TB     Complete            Done
    6       1TB     Dropped             Done
Next
    Local
    0       250GB   Documents -Backup   Done
    1       1TB     Hallmark 2          Done
    2       1TB     Hallmark            Done
